import { Component, OnInit } from '@angular/core';
import { EmployeeService } from '../employee.service';

@Component({
  selector: 'app-employee-list',
  template: `
  <h1 appStyleColor [color]="'white'" [bg]="'green'">Employee List</h1>
  <p>{{errMsg}}</p>
  <div *ngFor="let employee of employees">
  <h2>{{employee.id+"."+employee.name+" "+employee.model}}</h2>
  </div>
  `,
  styleUrls: ['./employee-list.component.css']
})
export class EmployeeListComponent implements OnInit {

  public employees=[];
  public errMsg;

  constructor(private employeeService : EmployeeService) { }

  ngOnInit() {
    this.employeeService.getEmployees().subscribe(data=>this.employees=data,error=>this.errMsg=error);
  }

}
