import { Injectable, Inject } from "@angular/core";
import { HttpClient, HttpHeaders } from "@angular/common/http";
//eslint-disable-next-line
import { Observable } from "rxjs";

@Injectable()
export default class SampleService {
  private _http: HttpClient;

  constructor(@Inject(HttpClient) http: HttpClient) { 
    this._http = http;
  }

  //eslint-disable-next-line
  public login(usuario: String, senha: String): boolean {
    return true;
  }

  //eslint-disable-next-line
  public send(cpf:string, user: string): Observable<Object> {
    let headers = new HttpHeaders();
    headers = headers.append("Content-Type", "application/json");
    headers = headers.append("User", user);
    return this._http.post(
      "https://autenticacao.xuxa.cloud.net/api/token",
      cpf,
      { headers: headers }
    );
  }
}
