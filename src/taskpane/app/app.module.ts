import { NgModule } from "@angular/core";
import { FormsModule, ReactiveFormsModule } from "@angular/forms";
import { BrowserModule } from "@angular/platform-browser";
import AppComponent from "./app.component";
import LoginComponent from "./components/login/login.component";
import { AppRoutingModule } from "./app.routing";
import { LocationStrategy, HashLocationStrategy } from "@angular/common";
import SampleService from "./services/sample.service";
import { HttpClientModule } from "@angular/common/http";
import ExcelAddinHelperService from "./components/excel-addin-helper/excel-addin-helper.service";
import MainComponent from "./components/main/main.component";

@NgModule({
  providers: [{ provide: LocationStrategy, useClass: HashLocationStrategy }, SampleService, ExcelAddinHelperService],
  declarations: [AppComponent, LoginComponent, MainComponent],
  imports: [BrowserModule, FormsModule, ReactiveFormsModule, AppRoutingModule, HttpClientModule],
  bootstrap: [AppComponent]
})
export default class AppModule {}
