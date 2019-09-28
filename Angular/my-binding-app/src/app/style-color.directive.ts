import { Directive, ElementRef, OnInit, Input, HostListener } from '@angular/core';

@Directive({
  selector: '[appStyleColor]'
})
export class StyleColorDirective implements OnInit{
  @Input() color : string;
  @Input() bg : string;
  constructor(private eleRef : ElementRef) { }
  @HostListener('mouseenter') onMouseEnter(){
    this.eleRef.nativeElement.style.color=this.color;
    this.eleRef.nativeElement.style.backgroundColor=this.bg;
  }
  @HostListener('mouseleave') onMouseLeave(){
    this.eleRef.nativeElement.style.color=null;
    this.eleRef.nativeElement.style.backgroundColor=null;
  }
  ngOnInit(){

  }

}
