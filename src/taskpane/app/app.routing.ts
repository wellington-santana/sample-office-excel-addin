/*eslint-disable */
import { NgModule } from "@angular/core";
import { RouterModule, Routes } from "@angular/router";
import LoginComponent from "./components/login/login.component";
import MainComponent from "./components/main/main.component";
/*eslint-enabled */
const routes: Routes = [
  { path: "login", component: LoginComponent },
  { path: "main", component: MainComponent },
  { path: '', redirectTo: '/login', pathMatch: 'full' },
];

@NgModule({
  exports: [RouterModule],
  imports: [RouterModule.forRoot(routes)]
})
export class AppRoutingModule {}
