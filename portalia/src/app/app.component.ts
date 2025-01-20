// src/app/app.component.ts
import { Component } from '@angular/core';

@Component({
  selector: 'app-root',
  standalone: false,
  template: `
    <h1>Welcome to {{ title }}</h1>
    <nav>
      <a [routerLink]="['/calculator']">Go to Calculator</a>
    </nav>
    <router-outlet></router-outlet>
  `
})
export class AppComponent {
  title = 'portalia';
}
