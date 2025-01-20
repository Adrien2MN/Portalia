import { Component } from '@angular/core';

@Component({
  selector: 'app-calculator',
  standalone: false, // Not a standalone component
  templateUrl: './calculator.component.html',
  styleUrls: ['./calculator.component.css']
})
export class CalculatorComponent {
  // Parameters for the form
  parameters = {
    calculationType: 'TJM', // Default: TJM
    contractType: 'CDI', // Default: CDI
    fraisGestion: 8, // Default frais de gestion
    provisions: 10, // Default provisions
    joursTravailles: 18, // Default jours travaillés
    ticketRestaurant: false // Default: no ticket restaurant
  };

  // Result object
  result: any = null;

  // Method to handle calculations
  calculate() {
    // Mock API call (replace this with an actual backend call later)
    this.result = {
      tjm: this.parameters.calculationType === 'TJM' ? 500 : null,
      brut: this.parameters.calculationType === 'BRUT' ? 7380 : null,
      net: this.parameters.calculationType === 'NET' ? 4870 : null
    };
  }
}
