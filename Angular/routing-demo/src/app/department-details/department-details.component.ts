import { Component, OnInit } from '@angular/core';
import {ActivatedRoute,Router,ParamMap} from '@angular/router';

@Component({
  selector: 'app-department-details',
  template: `
    <h2>
     Department id is {{departmentId}}
    </h2>
    <p>
    <a (click)="goPrevious()">back </a>
    <a (click)="goNext()">next</a><br>
    </p>
    <button (click)="goDepartments()">Back</button>
    <p>
    <button (click)="showOverview()">Overview</button>
    <button (click)="showContact()">Contact</button>
    </p>
    <router-outlet></router-outlet>
  `,
  styles: []
})
export class DepartmentDetailsComponent implements OnInit {
  public departmentId
  constructor(private route:ActivatedRoute,
    private router:Router
    ) { }

  ngOnInit() {
    //this.route.snapshot.paramMap.get('id')
    this.route.paramMap.subscribe((params : ParamMap )=> {
      let id=parseInt(params.get('id'));
      this.departmentId =id;
    });
    
  }

  goPrevious(){
    //this.router.navigate(['/department',this.departmentId-1])
    this.router.navigate([this.departmentId-1],{relativeTo:this.route})
  }

  goNext(){
    this.router.navigate([this.departmentId+1],{relativeTo:this.route})
  }

  goDepartments(){
    this.router.navigate(['../',{id:this.departmentId}],{relativeTo:this.route})
  }

  showOverview(){
    this.router.navigate(['overview'],{relativeTo:this.route})
  }

  showContact(){
    this.router.navigate(['contact'],{relativeTo:this.route})
  }
 
}
