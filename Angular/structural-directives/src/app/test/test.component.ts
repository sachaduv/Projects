import { Component, OnInit } from '@angular/core';

@Component({
  selector: 'app-test',
  template: `
  <div>
  <h2 *ngIf="displayName; then thenBlock; else elseBlock">If statement executed</h2>
  <ng-template #thenBlock><h2>Then block executed</h2></ng-template>
  <ng-template #elseBlock><h2>Else block executed</h2></ng-template>
  </div>
  <div [ngSwitch]="color">
  <h2 *ngSwitchCase="'red'">Red is selected</h2>
  <h2 *ngSwitchCase="'blue'">Blue is selected</h2>
  <h2 *ngSwitchCase="'Green'">Green is selected</h2>
  <h2 *ngSwitchDefault>pick either red , blue, green </h2>
  </div>
  <div *ngFor="let color of colors;index as i">
  <!-- odd as o ; first as f; last as l;even as e-->
   <h2>{{i}} {{color}}</h2>
  </div>
  `,
  styleUrls: ['./test.component.css']
})
export class TestComponent implements OnInit {
  displayName = true
  color="orange"
  colors = ['red','green','blue','yellow']
  constructor() { }

  ngOnInit() {
  }

}
