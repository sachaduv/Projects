import { Component, OnInit } from '@angular/core';
import {Router, ActivatedRoute, ParamMap} from '@angular/router';

@Component({
  selector: 'app-department-list',
  template: `
  <h2>Departments List</h2>
  <ul>
    <li *ngFor="let department of departments " [class.selected]="isSelected(department)">
     {{department.id}} <a (click)="onSelect(department) " >{{department.name}}</a>
    </li>
  </ul>
  `,
  styles : [`
  li.selected {
    color:red;
  }
  `]
})
export class DepartmentListComponent implements OnInit {
  public urlId;
  departments = [
    {"id":1,"name":'Angular'},
    {"id":2,"name":'React'},
    {"id":3,"name":'Node'},
    {"id":4,"name":'Express'}
  ]
  constructor(private router:Router,private route:ActivatedRoute) { 
  }

  ngOnInit() {
    this.route.paramMap.subscribe((param:ParamMap)=>{
      let id = parseInt(param.get('id'))
      this.urlId=id;
    })
  }

  onSelect(department){
    this.router.navigate([department.id],{relativeTo : this.route})
  }

  isSelected(department){
    return department.id === this.urlId
  }

}
