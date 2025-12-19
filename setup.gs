function manualSetup() {
  const url = "https://docs.google.com/spreadsheets/d/1J0bMQamssoKD9OFO5HLVampKYWhWUfaljlUY3O--7us/edit?gid=2019150533#gid=2019150533";
  const settings = { dataHubUrl: url };
  PropertiesService.getScriptProperties().setProperty('adminSettings', JSON.stringify(settings));
  console.log("Settings saved manually.");
}