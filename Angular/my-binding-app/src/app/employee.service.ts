import { Injectable } from '@angular/core';
import { HttpClient, HttpErrorResponse } from '@angular/common/http';
import { IEmployee } from './iemployee';
import { Observable,throwError } from 'rxjs';
import {catchError} from 'rxjs/operators';

@Injectable({
  providedIn: 'root'
})

export class EmployeeService {
  public _url = "/assets/data/employee.json";
  constructor(private http : HttpClient) { }

  errorHandler(error : HttpErrorResponse){
    return throwError(error.message||"Server Error");
  }

  getEmployees():Observable<IEmployee[]>{
    return this.http.get<IEmployee[]>(this._url).pipe(catchError(this.errorHandler));
  }

  
}
