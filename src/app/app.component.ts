import { Component } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { ExtractDataComponent } from './extract-data/extract-data.component';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet, ExtractDataComponent],
  templateUrl: './app.component.html',
  styleUrl: './app.component.css'
})
export class AppComponent {
  title = 'Angular Add-in';
}
