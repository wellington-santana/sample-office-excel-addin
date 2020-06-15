import { Component, Inject } from "@angular/core";
import { Router } from "@angular/router";
import SampleService from "../../services/sample.service";
const template = require("./login.component.html");
// const styles = require("./app.component.css");
/* global require */

@Component({
  selector: "login-component",
  template: template
})
export default class LoginComponent {
  router: Router;
  sampleService: SampleService;
  isLoading: boolean = false;
  isShowAlert: boolean = false;
  message: String;

  constructor(@Inject(Router) router: Router, @Inject(SampleService) sampleService: SampleService) {
    this.router = router;
    this.sampleService = sampleService;
  }

  showAlert(message: String) {
    this.isShowAlert = true;
    this.message = message;
  }

  login(usuario: String, senha: String) {
    this.isLoading = true;
    this.isShowAlert = false;

    let isAuth = this.sampleService.login(usuario, senha);
    if (isAuth) {
      this.router.navigate(["main"]);
    } else {
      this.isLoading = false;
      this.showAlert("Erro ao tentar realizar login, verifique as informações e tente novamente.");
    }
  }
}
