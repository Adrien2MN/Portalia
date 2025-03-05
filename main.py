from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import os
import tempfile
import shutil
from typing import Optional
import xlwings as xw
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = FastAPI()

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allows all origins. Replace with specific origins for production.
    allow_credentials=True,
    allow_methods=["*"],  # Allows all HTTP methods
    allow_headers=["*"],  # Allows all headers
)

# Path to the Excel template - using the new XLSM file
EXCEL_TEMPLATE_PATH = "PORTALIA MC2 CONSULTANTS 2024 V03-24.xlsm"

@app.get("/")
def read_root():
    return {"message": "Bienvenue sur FastAPI"}

@app.get("/convert")
async def convert(
    tjm: Optional[float] = None,
    jours_travailles: Optional[int] = None,
    contract_type: Optional[str] = None,
    frais_fonctionnement: Optional[float] = None,
    ticket_restaurant: Optional[bool] = None,
    mutuelle: Optional[bool] = None,
    code_commune: Optional[str] = None
):
    # Check if we have the required parameters
    if tjm is None or jours_travailles is None:
        raise HTTPException(status_code=400, detail="TJM and jours_travailles are required")
    
    try:
        logger.info(f"Starting calculation with TJM={tjm}, jours={jours_travailles}")
        
        # Create a temporary copy of the template
        temp_dir = tempfile.mkdtemp()
        temp_excel_path = os.path.join(temp_dir, "temp_calculation.xlsm")
        shutil.copy2(EXCEL_TEMPLATE_PATH, temp_excel_path)
        
        logger.info(f"Copied template to {temp_excel_path}")
        
        # Open the Excel file with xlwings
        app_excel = xw.App(visible=False)
        wb = app_excel.books.open(temp_excel_path)
        
        # Try to get the correct sheet name
        sheet_names = [sheet.name for sheet in wb.sheets]
        logger.info(f"Available sheets: {sheet_names}")
        
        # Look for the sheet that most likely contains calculation with provisions
        calculation_sheet = "1. Calcul Avec prov"
        logger.info(f"Using calculation sheet: {calculation_sheet}")
        ws = wb.sheets[calculation_sheet]
        
        # Fill in the data
        # You may need to adjust cell references based on the actual XLSM structure
        try:
            ws.range("J4").value = tjm
            logger.info(f"Set TJM to {tjm}")
            
            ws.range("J5").value = jours_travailles
            logger.info(f"Set jours travaillés to {jours_travailles}")
            
            # Handle contract type
            if contract_type == "CDI":
                ws.range("J8").value = "2%"
                ws.range("J9").value = "A négocier"
                ws.range("J10").value = "0%"
                logger.info("Set contract type to CDI")
            elif contract_type == "CDD":
                ws.range("J8").value = "0%"
                ws.range("J9").value = "0%"
                ws.range("J10").value = "10%"
                logger.info("Set contract type to CDD")
            
            # Handle frais de fonctionnement
            if frais_fonctionnement is not None:
                ws.range("J12").value = frais_fonctionnement
                logger.info(f"Set frais de fonctionnement to {frais_fonctionnement}")
            
            # Handle ticket restaurant
            if ticket_restaurant:
                ws.range("J21").value = 198
                logger.info("Enabled ticket restaurant")
            else:
                ws.range("J21").value = 0
                logger.info("Disabled ticket restaurant")
            
            # Handle mutuelle
            if mutuelle:
                ws.range("J17").value = "Oui"
                logger.info("Enabled mutuelle")
            else:
                ws.range("J17").value = "Non"
                logger.info("Disabled mutuelle")
            
            # Handle code commune
            if code_commune:
                ws.range("J25").value = code_commune
                logger.info(f"Set code commune to {code_commune}")
        
        except Exception as e:
            logger.error(f"Error setting values: {str(e)}")
            raise HTTPException(status_code=500, detail=f"Error setting Excel values: {str(e)}")
        
        # Find available macros
        try:
            macro_names = [m.name for m in wb.macro("ThisWorkbook").module.procedures]
            logger.info(f"Available macros: {macro_names}")
        except Exception as e:
            logger.warning(f"Could not retrieve macro names: {str(e)}")
            macro_names = []
        
        # Run the macro to update Template 3 if available
        update_macro_name = None
        for macro_name in macro_names:
            if "update" in macro_name.lower() and "template" in macro_name.lower():
                update_macro_name = macro_name
                break
        
        if update_macro_name:
            try:
                logger.info(f"Running macro: {update_macro_name}")
                wb.macro(update_macro_name)()
            except Exception as e:
                logger.error(f"Macro execution error: {str(e)}")
                # Continue even if macro fails
        else:
            logger.warning("No update template macro found. Calculations may not be complete.")
        
        # Find the template sheet
        template_sheet = None
        for sheet_name in sheet_names:
            if "template" in sheet_name.lower():
                template_sheet = sheet_name
                break
        
        if not template_sheet:
            template_sheet = sheet_names[-1]  # Use last sheet as fallback
        
        logger.info(f"Using template sheet: {template_sheet}")
        template3 = wb.sheets[template_sheet]
        
        # Try to locate the result cells - this may need adjustment
        # Log cell values for debugging
        cell_values = {}
        for row in range(5, 20):
            for col in ['B', 'C']:
                cell_ref = f"{col}{row}"
                cell_values[cell_ref] = template3.range(cell_ref).value
        
        logger.info(f"Cell values: {cell_values}")
        
        # Assuming the important results are in these cells (adjust as needed)
        result = {
            "tjm": tjm,
            "brut_mensuel": template3.range("C10").value,
            "net_mensuel": template3.range("C12").value,
            "frais_gestion": template3.range("C14").value,
            "autres_details": {
                "ticket_restaurant_contribution": template3.range("C16").value if ticket_restaurant else 0,
                "mutuelle_contribution": template3.range("C18").value if mutuelle else 0,
            }
        }
        
        logger.info(f"Calculation result: {result}")
        
        # Save and close the Excel file
        wb.save()
        wb.close()
        app_excel.quit()
        
        # Clean up temporary files
        shutil.rmtree(temp_dir)
        
        return result
        
    except Exception as e:
        logger.error(f"Excel processing error: {str(e)}")
        raise HTTPException(status_code=500, detail=f"Excel processing error: {str(e)}")

# Keep the old convert endpoint for backward compatibility
@app.get("/old_convert")
def old_convert(
    tjm: float = None, brut: float = None, net: float = None, 
    jours: int = 18, frais_fixes: float = 0.08, provisions: float = 0.10, 
    charges_sal: float = 0.22, charges_pat: float = 0.12
):
    if tjm:
        brut = 198 + (tjm * jours * (1 - frais_fixes - provisions))
        net = brut * (1 - charges_sal - charges_pat)
    elif brut:
        tjm = (brut - 198) / (jours * (1 - frais_fixes - provisions))
        net = brut * (1 - charges_sal - charges_pat)
    elif net:
        brut = net / (1 - charges_sal - charges_pat)
        tjm = (brut - 198) / (jours * (1 - frais_fixes - provisions))
    return {
    "tjm": round(tjm, 2) if tjm else None,
    "brut": round(brut, 2) if brut else None,
    "net": round(net, 2) if net else None
}