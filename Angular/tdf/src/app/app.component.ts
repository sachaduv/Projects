import { Component } from '@angular/core';
import {User} from './user';
import { EnrollementService } from './enrollement.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  constructor(private _enrollment:EnrollementService){}
  topics = ['Angular','React','Vue'];
  topicHasError=true;
  submitted=false;
  userModel = new User('Rob','rob@email.com',9848634039,'default','morning',true);
  errorMsg=''
  validateTopic(topic){
    if(topic==='default'){
      return this.topicHasError=true;
    }
    else{
     return this.topicHasError=false;
    }
  }
  onSubmit(regForm){
    console.log(regForm);
     this.submitted=true;
     this._enrollment.enrollUser(this.userModel).subscribe(
       data=>console.log('sucess',data),
      error=>this.errorMsg=error.statusText
    )
  }
}
