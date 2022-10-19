// 5e7baeb3-7e6a-4a17-8893-65f5547de834
function launchOneDrivePicker(){
    var odOptions = {
        clientId: "18f24b04-fc0b-49b9-9e60-b797818d090e",
        action: "download",
        multiSelect: true,
        advanced: {
        redirectUri:"http://localhost:4200",
        filter:"folder,.pptx,.jpeg,.jpg",
        },
        
        viewType: "all",
        
        success: function(files) {
          OneDriveFilesReference.sendOneDriveFiles(files);
        },
        cancel: function() { console.log("cancel") },
        error: function() { console.log("error") },
      }
    return OneDrive.open(odOptions);
  }

var OneDriveFilesReference = null;