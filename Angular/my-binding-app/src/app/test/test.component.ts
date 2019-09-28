import { Component, OnInit, Input, Output,EventEmitter } from '@angular/core';
import { InjectableCompiler } from '@angular/compiler/src/injectable_compiler';

@Component({
  selector: 'app-test',
  template: `
  <div>
  <h2>Welcome {{name}}</h2>
  <h2>{{getName()}}</h2>
  <input bind-value="name" type="text">
  <input [disabled]="isDisabled" value="Naruto" type="text">
  </div>
  <div>
  <p class="text-sucess">Congratulations</p>
  <p [class]="sucessClass">Congrats</p>
  <p [class.text-danger]="hasError">Errored Out</p>
  <p [ngClass]="displayClass">Hello this ngClass Styles</p>
  </div>
  <div>
  <h2 [style.color]="hasError ? 'red' : 'green'"> Style Binding </h2>
  <h2 [style.color]="styleColor"> Style Color</h2>
  <h2 [ngStyle]="myStyles">Hello this is Angular's ngStyle</h2>
  </div>
  <div>
  <button (click)="onClick($event)">event Binding</button>
  <button (click)="greetings='Welcome to Angulars Event Binding'">click!..</button>
  {{greetings}}
  </div>
  <div>
  <input #log type="text">
  <button (click)="logger(log.value)">log</button>
  </div>
  <div>
  <input [(ngModel)]="name" type="text" placeholder="ngModel">
  {{name}}
  </div>
  <div>
  {{fromParent}}
  </div>
  <div>
  <button (click)="toAppComponent()">@Output</button>
  </div>
  <div>
  <p>{{"Pipes " + name | lowercase}}</p>
  <p>{{name | uppercase}}</p>
  <p>{{message | titlecase}}</p>
  <p>{{displayClass | json}}</p>
  <p>{{message | slice : 2 : 6}}

  <p>{{5.23 | number : '3.5-7'}}</p>
  
  <p>{{1238 | currency }}</p>
  <p>{{1234 | currency : 'GBP'}}</p>
  <p>{{3245 | currency : 'EUR' : 'code'}}</p>

  <p>{{date}}</p>
  <p>{{date | date:'short'}}</p>
  <p>{{date | date:'shortDate'}}</p>
  <p>{{date | date:'shortTime'}}</p>
  </div>
  `,
  styles: [`
  .text-sucess{
    color : green;
  }
  .text-danger{
    color : red
  }
  .text-special{
    font-style:italic;
  }
  `
  ]
  
  
})
export class TestComponent implements OnInit {
  public date = new Date()
  public name = 'Hinata'
  public message = "hElllo , whO is tHis"
  public sucessClass="text-sucess"
  public hasError=false
  public isSpecial=true
  public isDisabled = true
  public styleColor = 'blue'
  public greetings=''
  public displayClass = {
       "text-sucess":!this.hasError,
       "text-danger":this.hasError,
       "text-special":this.isSpecial
  }
  public myStyles={
    "color":"orange",
    "fontStyle":"italic"
  }
  @Input('parentData') public fromParent;
  @Output() public childEvent = new EventEmitter()
  constructor() { }

  ngOnInit() {
  }

  getName():string{
    return "Hello getName is calling "+this.name
  }

  onClick(event):void{
    console.log(event)
    this.greetings = "Hello, this is Event Binding"
  }

  logger(logData):void{
    console.log(logData)
  }
  
  toAppComponent():void{
    this.childEvent.emit('send data to parent component from child component')
  }
}
