import { Component, OnInit } from '@angular/core';
import { EmployeeService } from '../employee.service';

@Component({
  selector: 'app-employee-detail',
  template:  `
  <h1 appStyleColor [color]="'red'" [bg]="'yellow'">Employee Details</h1>
  <p>{{errMsg}}</p>
  <div *ngFor="let employee of employees">
  <h2>{{employee.id + ":" + employee.name+"-"+employee.model}}</h2>
  </div>
  `
  ,
  styleUrls: ['./employee-detail.component.css']
})
export class EmployeeDetailComponent implements OnInit {

  public employees = [];
  public errMsg;

  constructor(private employeeService:EmployeeService) { }

  ngOnInit() {
    this.employeeService.getEmployees().subscribe(data => this.employees=data,error=>this.errMsg=error);
  }

}
