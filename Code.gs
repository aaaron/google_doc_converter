
function main() {
  // Log the name of every file in the folder.
  folderName = "ROOT_FOLDER_NAME"
  Logger.log("Starting with folder " + folderName);
  var folders = DriveApp.getFoldersByName(folderName);
  
  while(folders.hasNext()) {
    var folder = folders.next();
    if (folders.hasNext()) {
      Logger.log("Too many matching folders!");
      break;
    }
    convertFolder(folder, 1);
  }
  Logger.log("Done.");
}

function convertFolder(folder, depth)
{
  types_to_convert = ["application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                      "application/msword",
                      "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                      "application/vnd.ms-powerpoint",
                      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                      "application/vnd.ms-excel"]
  
  // skip any folder with archive in the name
  if(folder.getName().toLowerCase().indexOf("archive") != -1) {
    Logger.log("Skipping folder " + folder);
    return;
  } 
  
  Logger.log(depth + ": Converting folder " + folder);
  // convert files
  for(var i = 0; i < types_to_convert.length; i++) {
    // get mimeType
    mimeType = types_to_convert[i];
    var files = folder.getFilesByType(mimeType);
    
    // list files
    while (files.hasNext()) {
      var file = files.next();
      fileName = file.getName();
      // skip semi-restored msft files; ie, files with tildes
      if(fileName.indexOf("~") == -1)
      {
        var fileHash = {
          title: fileName,
          mimeType: file.getMimeType(),
          parents: [{"id": folder.getId()}]
        };
        
        Logger.log(depth + ": Converting file " + fileName);
        try {
          Drive.Files.insert(fileHash, file.getBlob(), {convert : true});
          file.setTrashed(true);
        } catch(e) {
          // if can't convert a file, put tilde at end
          file.setName(fileName + "~");
          Logger.log("Error " + e);
        }
      }
      
    }
  }

  // recurse folders
  var folders = folder.getFolders();
  while (folders.hasNext()) {
     convertFolder(folders.next(), depth+1); // recurse
  }
}
