import { Component } from '@angular/core';
import { HttpClient } from '@angular/common/http';

@Component({
  selector: 'app-calculator',
  standalone: false, // Not a standalone component
  templateUrl: './calculator.component.html',
  styleUrls: ['./calculator.component.css']
})
export class CalculatorComponent {
  // Parameters for the form
  parameters = {
    calculationType: 'TJM', // Default calculation type
    tjm: 0,
    brut: 0,
    net: 0,
    joursTravailles: 18, // Default: 18 days
    fraisGestion: 8, // Default: 8%
    provisions: 10, // Default: 10%
    ticketRestaurant: false,
    contractType: 'CDI'
  };
  
  // Result object
  result: any = null;
  // Error message
  errorMessage: string = '';
  
  constructor(private http: HttpClient) {}
  
  // Call the backend API
  calculate() {
    // Reset error message
    this.errorMessage = '';
    
    const queryParams: { [key: string]: string } = {
      tjm: this.parameters.tjm !== null ? this.parameters.tjm.toString() : '',
      brut: this.parameters.brut !== null ? this.parameters.brut.toString() : '',
      net: this.parameters.net !== null ? this.parameters.net.toString() : '',
      jours: this.parameters.joursTravailles.toString(),
      frais_fixes: (this.parameters.fraisGestion / 100).toString(),
      provisions: (this.parameters.provisions / 100).toString(),
      charges_sal: '0.22', // Default charges
      charges_pat: '0.12'  // Default charges
    };
    
    // Filter out empty parameters
    const filteredQueryParams = Object.fromEntries(
      Object.entries(queryParams).filter(([_, value]) => value !== '')
    );
    
    const queryString = new URLSearchParams(filteredQueryParams).toString();
    const apiUrl = `http://127.0.0.1:8000/convert?${queryString}`;
    
    this.http.get(apiUrl).subscribe(
      (response) => {
        this.result = response; // Store the result
        // Scroll to results
        setTimeout(() => {
          const resultElement = document.querySelector('.result-container');
          if (resultElement) {
            resultElement.scrollIntoView({ behavior: 'smooth' });
          }
        }, 100);
      },
      (error) => {
        console.error('Error calling backend API:', error);
        this.errorMessage = 'Une erreur est survenue lors du calcul. Veuillez réessayer.';
      }
    );
  }
}