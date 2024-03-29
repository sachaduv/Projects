import { Component, OnInit,Input } from '@angular/core';
import { Location } from '@angular/common';
import { ActivatedRoute } from '@angular/router'; 

import {Hero} from '../hero';
import { HeroService } from '../hero.service';

@Component({
  selector: 'app-hero-detail',
  templateUrl: './hero-detail.component.html',
  styleUrls: ['./hero-detail.component.css']
})
export class HeroDetailComponent implements OnInit {

  hero:Hero;
  constructor(
    private herService:HeroService,
    private location:Location,
    private route:ActivatedRoute
  ) { }

  ngOnInit() {
    this.getHero();
  }

  getHero():void{
    const id=+this.route.snapshot.paramMap.get('id');
    this.herService.getHero(id).subscribe(hero=>this.hero=hero);
  }
  goBack():void{
    this.location.back();
  }

}
