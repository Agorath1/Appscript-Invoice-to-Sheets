function import_xls(sheetID) {
  const file = DriveApp.getFileById(sheetID);
  const file_name = file.getName();

  if (file_name.match(/\.(xls|xlsx)$/)){
    const blob = file.getBlob();
    const resource = {
      name: file_name,
      parents: file.getParents(),
      mimeType: MimeType.GOOGLE_SHEETS
    };

    try {
      const converted_file = Drive.Files.copy(resource, sheetID, {convert: true});
      Logger.log("Converted file ID:" + sheetID);

      Drive.Files.remove(sheetID)
    } catch (e) {
      Logger.log("Error converting file ${file_name}: ${e.message}");
    }
  }
}
