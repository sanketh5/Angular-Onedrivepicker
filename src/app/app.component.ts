import { Component } from '@angular/core';
declare function launchOneDrivePicker(): any;
declare var OneDriveFilesReference: any;

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  title = 'onedrive-picker';


  ngOnInit() {
    OneDriveFilesReference = this;
  }

  ngOnDestroy(): any {
    OneDriveFilesReference = null;
  }

  
  sendOneDriveFiles(files: any) {
    /* Creating a list of links to the files that were selected in the One Drive picker. */
    files = files['value']
    console.warn(files['value'])
    for (let file of files){
      let ul = document.getElementById("results");
      let li = document.createElement("li");
      let a = document.createElement('a');
      a.setAttribute('href',file['@microsoft.graph.downloadUrl']);
      a.setAttribute('target','_blank');
      a.innerHTML = "Click here";
      li.appendChild(a);
      ul?.appendChild(li);
  }
  }

  loadOneDrive() {
    launchOneDrivePicker();
  }
}
