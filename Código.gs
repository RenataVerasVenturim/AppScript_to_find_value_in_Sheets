function doGet(e) {
  var template=HtmlService.createTemplateFromFile("Principal");
  return template.evaluate()
    .setTitle('UFF')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .addMetaTag('viewport','width=device-width,initial-scale=1');
}


function Chamar(Arquivo){

  return HtmlService.createHtmlOutputFromFile(Arquivo).getContent();
}



