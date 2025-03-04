import { Component, OnInit } from '@angular/core';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';

interface CalculationResult {
  tjm: number;
  brut_mensuel: number;
  net_mensuel: number;
  frais_gestion: number;
  autres_details: {
    ticket_restaurant_contribution: number;
    mutuelle_contribution: number;
  };
}

@Component({
  selector: 'app-calculator',
  standalone: false,
  templateUrl: './calculator.component.html',
  styleUrls: ['./calculator.component.css']
})
export class CalculatorComponent implements OnInit {
  // Parameters for the form
  parameters = {
    tjm: 500, // Default TJM value
    joursTravailles: 18, // Default: 18 days
    contractType: 'CDI', // Default: CDI
    fraisFonctionnement: 8, // Default: 8%
    ticketRestaurant: false,
    mutuelle: false,
    codeCommune: ''
  };

  // Result object
  result: CalculationResult | null = null;
  isLoading: boolean = false;
  errorMessage: string | null = null;
  debugMode: boolean = false; // Set to true to see raw API response
  
  constructor(private http: HttpClient) {}
  
  ngOnInit(): void {
    // You could load saved preferences here if needed
  }

  // Format currency values for display
  formatCurrency(value: any): string {
    if (value === null || value === undefined) {
      return 'N/A';
    }
    
    // Handle various formats that might come from Excel
    let numValue: number;
    
    if (typeof value === 'string') {
      // Remove any currency symbols or spaces
      const cleanValue = value.replace(/[^0-9.,]/g, '').replace(',', '.');
      numValue = parseFloat(cleanValue);
    } else {
      numValue = Number(value);
    }
    
    if (isNaN(numValue)) {
      return 'N/A';
    }
    
    return numValue.toLocaleString('fr-FR', { style: 'currency', currency: 'EUR' });
  }
  
  // Toggle debug mode
  toggleDebugMode(): void {
    this.debugMode = !this.debugMode;
  }

  // Call the backend API
  calculate(): void {
    // Validate inputs before making API call
    if (this.parameters.tjm <= 0) {
      this.errorMessage = "Le taux journalier doit être supérieur à 0.";
      return;
    }
    
    if (this.parameters.joursTravailles <= 0 || this.parameters.joursTravailles > 30) {
      this.errorMessage = "Le nombre de jours travaillés doit être entre 1 et 30.";
      return;
    }
    
    this.errorMessage = null;
    this.isLoading = true;
    
    const queryParams: { [key: string]: string } = {
      tjm: this.parameters.tjm.toString(),
      jours_travailles: this.parameters.joursTravailles.toString(),
      contract_type: this.parameters.contractType,
      frais_fonctionnement: (this.parameters.fraisFonctionnement / 100).toString(),
      ticket_restaurant: this.parameters.ticketRestaurant.toString(),
      mutuelle: this.parameters.mutuelle.toString(),
      code_commune: this.parameters.codeCommune
    };

    // Create the query string
    const queryString = new URLSearchParams(queryParams).toString();
    const apiUrl = `http://127.0.0.1:8000/convert?${queryString}`;

    this.http.get<CalculationResult>(apiUrl).subscribe({
      next: (response: CalculationResult) => {
        this.result = response; // Store the result
        this.isLoading = false;
        console.log('Calculation result:', this.result);
        
        // Handle null or undefined values
        if (!this.result) {
          this.errorMessage = "Le calcul n'a pas généré de résultats valides.";
        } else if (!this.result.brut_mensuel && !this.result.net_mensuel) {
          this.errorMessage = "Les résultats de salaire sont incomplets ou invalides.";
        }
      },
      error: (error: HttpErrorResponse) => {
        this.isLoading = false;
        this.errorMessage = "Une erreur s'est produite lors de la communication avec le serveur. Veuillez réessayer.";
        console.error('Error calling backend API:', error);
        
        // Include error details in the error message if available
        if (error.error && error.error.detail) {
          this.errorMessage += ` Détail: ${error.error.detail}`;
        }
      }
    });
  }
  
  // Reset the form to defaults
  resetForm(): void {
    this.parameters = {
      tjm: 500,
      joursTravailles: 18,
      contractType: 'CDI',
      fraisFonctionnement: 8,
      ticketRestaurant: false,
      mutuelle: false,
      codeCommune: ''
    };
    this.result = null;
    this.errorMessage = null;
  }
}