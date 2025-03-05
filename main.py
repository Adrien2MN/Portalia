from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
import os
import tempfile
import shutil
from typing import Optional
import logging

# Configure logging
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

app = FastAPI()

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allows all origins
    allow_credentials=True,
    allow_methods=["*"],  # Allows all HTTP methods
    allow_headers=["*"],  # Allows all headers
)

# Path to the Excel template
EXCEL_TEMPLATE_PATH = "PORTALIA MC2 CONSULTANTS 2024 V0324.xlsm"

@app.get("/")
def read_root():
    return {"message": "Bienvenue sur FastAPI"}

def str_to_bool(value: str) -> bool:
    """Convert string to boolean, handling various formats."""
    if not value:
        return False
    return value.lower() in ('true', 't', 'yes', 'y', '1')

@app.get("/convert")
async def convert(
    tjm: Optional[float] = Query(None),
    jours_travailles: Optional[int] = Query(None),
    contract_type: Optional[str] = Query(None),
    frais_fonctionnement: Optional[float] = Query(None),
    ticket_restaurant: Optional[str] = Query(None),
    mutuelle: Optional[str] = Query(None),
    code_commune: Optional[str] = Query(None)
):
    # Log the received parameters
    logger.info(f"Received parameters: tjm={tjm}, jours_travailles={jours_travailles}, " +
                f"contract_type={contract_type}, frais_fonctionnement={frais_fonctionnement}, " +
                f"ticket_restaurant={ticket_restaurant}, mutuelle={mutuelle}, code_commune={code_commune}")
    
    # Convert string boolean parameters to actual booleans
    ticket_restaurant_bool = str_to_bool(ticket_restaurant) if ticket_restaurant is not None else False
    mutuelle_bool = str_to_bool(mutuelle) if mutuelle is not None else False
    
    # Check if we have the required parameters
    if tjm is None or jours_travailles is None:
        error_msg = "TJM and jours_travailles are required"
        logger.error(error_msg)
        raise HTTPException(status_code=400, detail=error_msg)
    
    # Check if Excel file exists
    if not os.path.exists(EXCEL_TEMPLATE_PATH):
        error_msg = f"Excel template file not found: {EXCEL_TEMPLATE_PATH}"
        logger.error(error_msg)
        raise HTTPException(status_code=500, detail=error_msg)
    
    # Import xlwings here to avoid startup errors if Excel is not available
    try:
        import xlwings as xw
    except ImportError:
        error_msg = "xlwings module not installed. Please install it with: pip install xlwings"
        logger.error(error_msg)
        raise HTTPException(status_code=500, detail=error_msg)
    
    try:
        logger.info(f"Starting Excel calculation with TJM={tjm}, jours={jours_travailles}")
        
        # Create a temporary copy of the template
        temp_dir = tempfile.mkdtemp()
        temp_excel_path = os.path.join(temp_dir, "temp_calculation.xlsm")
        shutil.copy2(EXCEL_TEMPLATE_PATH, temp_excel_path)
        
        logger.info(f"Copied template to {temp_excel_path}")
        
        # Open the Excel file with xlwings
        app_excel = xw.App(visible=False)
        wb = app_excel.books.open(temp_excel_path)
        
        try:
            # Try to get the correct sheet name
            sheet_names = [sheet.name for sheet in wb.sheets]
            logger.info(f"Available sheets: {sheet_names}")
            
            # Look for the sheet that most likely contains calculation with provisions
            calculation_sheet = "1. Calcul Avec prov"
            
            logger.info(f"Using calculation sheet: {calculation_sheet}")
            ws = wb.sheets["1. Calcul Avec prov"]
            
            # Fill in the data
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
            if ticket_restaurant_bool:
                ws.range("J21").value = 198
                logger.info("Enabled ticket restaurant")
            else:
                ws.range("J21").value = 0
                logger.info("Disabled ticket restaurant")
            
            # Handle mutuelle
            if mutuelle_bool:
                ws.range("J17").value = "Oui"
                logger.info("Enabled mutuelle")
            else:
                ws.range("J17").value = "Non"
                logger.info("Disabled mutuelle")
            
            # Handle code commune
            if code_commune:
                ws.range("J25").value = code_commune
                logger.info(f"Set code commune to {code_commune}")
                
            # Find available macros
            
            # Replace the macro listing section with this more robust code
            try:
                # First try to access ThisWorkbook module
                macro_names = []
                try:
                    macro_names = [m.name for m in wb.macro("ThisWorkbook").module.procedures]
                    logger.info(f"Found macros in ThisWorkbook: {macro_names}")
                except AttributeError:
                    # If that fails, try to find macros in other modules
                    logger.warning("Could not access ThisWorkbook.module, trying alternative methods")
                    try:
                        for module_name in wb.api.VBProject.VBComponents:
                            try:
                                module = wb.macro(module_name.Name)
                                if hasattr(module, 'module') and hasattr(module.module, 'procedures'):
                                    module_macros = [m.name for m in module.module.procedures]
                                    logger.info(f"Found macros in module {module_name.Name}: {module_macros}")
                                    macro_names.extend(module_macros)
                            except Exception as e:
                                logger.warning(f"Could not access macros in module {module_name.Name}: {str(e)}")
                    except Exception as e:
                        logger.warning(f"Could not access VBProject: {str(e)}")
                
                # If we found macro names, look for the update template macro
                update_macro_name = None
                if macro_names:
                    for macro_name in macro_names:
                        if "update" in macro_name.lower() and "template" in macro_name.lower():
                            update_macro_name = macro_name
                            break
                
                # If we still don't have a macro name, try some common names
                if not update_macro_name:
                    common_names = ["UpdateTemplate", "UpdateTemplate3", "Calculate", "Update", "CalculateTemplate"]
                    for name in common_names:
                        try:
                            # Try to run the macro directly without checking if it exists
                            wb.macro(name)()
                            logger.info(f"Successfully ran macro: {name}")
                            update_macro_name = name
                            break
                        except Exception as e:
                            logger.warning(f"Failed to run macro {name}: {str(e)}")
                
                # If we found a macro, run it
                if update_macro_name:
                    logger.info(f"Running macro: {update_macro_name}")
                    wb.macro(update_macro_name)()
                else:
                    # If we can't find any macros, just continue - the Excel file might calculate automatically
                    logger.warning("No macros found or runnable. Continuing without running macros.")
                    
            except Exception as e:
                logger.error(f"Error with Excel macros: {str(e)}")
                # Continue without failing - we might still get results
    
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
            
            # Get the results from the template
            result = {
                "tjm": tjm,
                "brut_mensuel": template3.range("C10").value,
                "net_mensuel": template3.range("C12").value,
                "frais_gestion": template3.range("C14").value,
                "autres_details": {
                    "ticket_restaurant_contribution": template3.range("C16").value if ticket_restaurant_bool else 0,
                    "mutuelle_contribution": template3.range("C18").value if mutuelle_bool else 0,
                }
            }
            
            logger.info(f"Excel calculation result: {result}")
            return result
        
        finally:
            # Make sure we always close Excel and clean up
            try:
                wb.save()
                wb.close()
                app_excel.quit()
                shutil.rmtree(temp_dir)
            except Exception as e:
                logger.error(f"Error cleaning up Excel: {str(e)}")
    
    except Exception as e:
        error_msg = f"Excel processing error: {str(e)}"
        logger.error(error_msg)
        raise HTTPException(status_code=500, detail=error_msg)

# Keep the old endpoint for backward compatibility
@app.get("/old_convert")
def old_convert(
    tjm: float = None, brut: float = None, net: float = None, 
    jours: int = 18, frais_fixes: float = 0.08, provisions: float = 0.10, 
    charges_sal: float = 0.22, charges_pat: float = 0.12
):
    # Redirect to the new endpoint with appropriate parameters
    return {"message": "This endpoint is deprecated. Please use /convert instead."}