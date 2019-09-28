import { Component, OnInit } from '@angular/core';
//import {FormGroup,FormControl} from '@angular/forms';
import {FormBuilder,Validators, FormGroup,FormArray } from '@angular/forms';
import {forbiddenNameValidator} from './shared/username-validator';
import {passwordValidator} from './shared/password-validator';
import { RegistrationService } from './registration.service';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
}) 
export class AppComponent implements OnInit {
  registrationForm : FormGroup;
  errorMsg='';
  submitted = false;
  constructor(private fb : FormBuilder,private registerService : RegistrationService){
  }
  ngOnInit(){
    this.registrationForm=this.fb.group({
      username : ['',[Validators.required,Validators.minLength(3),forbiddenNameValidator(/password/)]],
      email :[''],
      subscribe : [false],
      password : [''],
      confirmPassword : [''],
      address : this.fb.group({
        city :[''],
        state : [''],
        postalCode : ['']
      }),
      alternateEmails : this.fb.array([]),
    },{validator : passwordValidator});
   this.registrationForm.get('subscribe').valueChanges.subscribe(checkedValue => {
     const email = this.registrationForm.get('email');
     if(checkedValue){
       email.setValidators(Validators.required);
     }
     else{
       email.clearValidators();
     }
     email.updateValueAndValidity();
   })
  }
  get username(){
    return this.registrationForm.get('username');
  }
  get email(){
    return this.registrationForm.get('email');
  }

  get alternateEmails(){
    return this.registrationForm.get('alternateEmails') as FormArray;
  }

  addAlternateEmail(){
    return this.alternateEmails.push(this.fb.control(''));
  }
  title = 'reactive-forms';

  
  // registrationForm = new FormGroup({
  //   username : new FormControl(''),
  //   password : new FormControl(''),
  //   confirmPassword : new FormControl(''),
  //   address : new FormGroup({
  //    city : new FormControl(''),
  //    state : new FormControl(''),
  //    postalCode : new FormControl('')
  //   })
  // })

  loadApiData(){
    //patchValue -- load partial values of a form , however the setValue we need to pass all values of a form
    this.registrationForm.setValue({
      username : 'Bobb',
      email : 'email@email.com',
      subscribe : true,
      password : 'Sai@9848',
      confirmPassword : 'Sai@9848',
      address : {
        city : 'Vskp',
        state : 'AP',
        postalCode : '530040'
      }
    })
  }

  onSubmit(){
    //console.log(this.registrationForm.value);
    this.submitted=true;
    this.registerService.register(this.registrationForm.value).subscribe(data=>console.log("sucess",data),error=>this.errorMsg=error.statusText);
  }
}
