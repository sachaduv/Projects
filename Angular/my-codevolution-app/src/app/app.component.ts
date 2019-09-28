import { Component } from '@angular/core';
import { Heroes } from './heroes';
import { from } from 'rxjs';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  public heroes : Heroes[]=[
    {id:1,name:'Iron Man'},{id:2,name:'Hulk'},{id:3,name:'Captain America'},{id:3,name:'Thor'}]
    title = "Code evolution";
}
