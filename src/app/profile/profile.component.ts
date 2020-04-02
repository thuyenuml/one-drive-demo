import { Component, OnInit } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { HttpClient } from '@angular/common/http';
declare var OneDrive: any;

const GRAPH_ENDPOINT = 'https://graph.microsoft.com/v1.0/me';

@Component({
  selector: 'app-profile',
  templateUrl: './profile.component.html',
  styleUrls: ['./profile.component.css']
})
export class ProfileComponent implements OnInit {
  profile;

  constructor(private authService: MsalService, private http: HttpClient) { }

  ngOnInit() {
    this.getProfile();
    this.authService.acquireTokenPopup({scopes: ["user.read"]});
  }

  getProfile() {
    this.http.get(GRAPH_ENDPOINT)
      .subscribe(profile => {
        this.profile = profile;
        console.info(this.profile);
      });
  }

  launchOneDrivePicker() {
    var odOptions = {
      clientId: "6226576d-37e9-49eb-b201-ec1eeb0029b6",
      action: "query",
      multiSelect: false,
      // scopes: "user.read files.read files.read.all sites.read.all",
      advanced: {
        // redirectUri: "http://localhost:4200/profile"
        queryParameters: "select=id,@microsoft.graph.downloadUrl",
        filter: "folder,.twdx" /* display folder and files with extension '.twdx' only */
      },
      success: function(files) {
        // let getUrl = '@microsoft.graph.downloadUrl'; 
        console.info(Object.values(files.value[0])[1]);
        // const options = {
        //   authProvider,
        // };
        
        // const client = Client.init(options);
        
        // let res = await client.api(`/me/drive/items/${files.value[0].id}/content`)
          // .get();
        debugger 
      },
      cancel: function() { /* cancel handler */ },
      error: function(error) { /* error handler */ }
    }
    OneDrive.open(odOptions);  
}

}
