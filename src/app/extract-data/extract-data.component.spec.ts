import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ExtractDataComponent } from './extract-data.component';

describe('ExtractDataComponent', () => {
  let component: ExtractDataComponent;
  let fixture: ComponentFixture<ExtractDataComponent>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ExtractDataComponent]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ExtractDataComponent);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
