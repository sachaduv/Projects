import { Pipe,PipeTransform } from '@angular/core';
@Pipe({name:'heroFilter'})
export class HeroesFilter implements PipeTransform {
   transform(value:string,id:number):string{
       var     fav : string="";
        if(value=='Iron Man')
        {
            fav = 'is my favorate';
        }
        else{
            fav = "is not my fan"
        }
        return value +" Ranks "+ id.toString()+"("+fav+")";
    }   
}
