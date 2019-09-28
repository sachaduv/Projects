import { TestBed } from '@angular/core/testing';

import { EnrollementService } from './enrollement.service';

describe('EnrollementService', () => {
  beforeEach(() => TestBed.configureTestingModule({}));

  it('should be created', () => {
    const service: EnrollementService = TestBed.get(EnrollementService);
    expect(service).toBeTruthy();
  });
});
