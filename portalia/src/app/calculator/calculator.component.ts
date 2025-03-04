import { Component } from '@angular/core';
import { HttpClient } from '@angular/common/http';

@Component({
  selector: 'app-calculator',
  standalone: false,  // Explicitly set to false
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
  
  constructor(private http: HttpClient) {}

  // Call the backend API
  calculate() {
    this.http.post('http://127.0.0.1:8000/calculate', this.parameters).subscribe(
      (response: any) => {
        this.result = response;
      },
      (error) => {
        console.error('Erreur API:', error);
      }
    );
  }
  
}