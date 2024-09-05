var estilos_sheet = PropertiesService.getDocumentProperties();

function onOpen() {
  SpreadsheetApp.getUi().createMenu('DyR')
    .addItem('Mostrar barra lateral', 'mostrarBarraLateral')
    .addToUi();
  
}
function mostrarBarraLateral(){
  var barra = HtmlService.createTemplateFromFile('BarraLateral')
  .evaluate()
  .setTitle('Barra lateral DyR');
  SpreadsheetApp.getUi().showSidebar(barra);
  
}

function aplicarEstilo(estilo){
    borrarEstilos();
    var hojaActual = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var celdas =hojaActual.getActiveRange();
    celdas.setFontColor(estilos_sheet.getProperty('colorLetra'+estilo))
        .setBackground(estilos_sheet.getProperty('colorFondo'+estilo))
        .setFontSize(estilos_sheet.getProperty('sizeFuente'+estilo))
        .setValue('Estilo '+estilo);
    
    if (estilos_sheet.getProperty('BordeSupCO'+estilo)!=null && estilos_sheet.getProperty('BordeSupST'+estilo)!=null){

      celdas.setBorder(true,null,null,null,null,true,estilos_sheet.getProperty('BordeSupCO'+estilo),obtenerEnumBorde(estilos_sheet.getProperty('BordeSupST'+estilo)));

    }
    
    if (estilos_sheet.getProperty('BordeIzqCO'+estilo)!=null && estilos_sheet.getProperty('BordeIzqST'+estilo)!=null){
      celdas.setBorder(null,null,true,null,true,null,estilos_sheet.getProperty('BordeIzqCO'+estilo),obtenerEnumBorde(estilos_sheet.getProperty('BordeIzqST'+estilo)));
    }
    if (estilos_sheet.getProperty('BordeInfCO'+estilo)!=null && estilos_sheet.getProperty('BordeInfST'+estilo)!=null){
      celdas.setBorder(null,true,null,null,null,true,estilos_sheet.getProperty('BordeInfCO'+estilo),obtenerEnumBorde(estilos_sheet.getProperty('BordeInfST'+estilo)));
    }
    if (estilos_sheet.getProperty('BordeDerCO'+estilo)!=null && estilos_sheet.getProperty('BordeDerST'+estilo)!=null){
      celdas.setBorder(null,null,null,true,true,null,estilos_sheet.getProperty('BordeDerCO'+estilo),obtenerEnumBorde(estilos_sheet.getProperty('BordeDerST'+estilo)));
    }
    
    

}

function obtenerEnumBorde(tipoBorde){
  switch(tipoBorde){
    case 'DOTTED' : return SpreadsheetApp.BorderStyle.DOTTED;
    case 'DASHED' : return SpreadsheetApp.BorderStyle.DASHED;
    case 'SOLID' : return SpreadsheetApp.BorderStyle.SOLID;
    case 'SOLID_MEDIUM' : return SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
    case 'SOLID_THICK' : return SpreadsheetApp.BorderStyle.SOLID_THICK;
    case 'DOUBLE' : return SpreadsheetApp.BorderStyle.DOUBLE;
    default: return null;

  }
}
function guardarEstilo(estilo){
    eliminarEstilo(estilo);
    
    var celda =SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();
    
    var bordes = celda.getBorder();
    if(bordes!=null){
      var borde_sup=bordes.getTop();
      var borde_inf=bordes.getBottom();
      var borde_der=bordes.getRight();
      var borde_izq=bordes.getLeft();

      if(borde_sup.getColor()!=null && borde_sup.getBorderStyle()!=null)
      {
        estilos_sheet.setProperty('BordeSupCO'+estilo, borde_sup.getColor().asRgbColor().asHexString());
        estilos_sheet.setProperty('BordeSupST'+estilo, borde_sup.getBorderStyle());
      }
      if(borde_inf.getColor()!=null && borde_inf.getBorderStyle()!=null)
      {
        estilos_sheet.setProperty('BordeInfCO'+estilo, borde_inf.getColor().asRgbColor().asHexString());
        estilos_sheet.setProperty('BordeInfST'+estilo, borde_inf.getBorderStyle());
      }
      if(borde_izq.getColor()!=null && borde_izq.getBorderStyle()!=null)
      {
        estilos_sheet.setProperty('BordeIzqCO'+estilo, borde_izq.getColor().asRgbColor().asHexString());
        estilos_sheet.setProperty('BordeIzqST'+estilo, borde_izq.getBorderStyle());
      }
      if(borde_der.getColor()!=null && borde_der.getBorderStyle()!=null)
      {
        estilos_sheet.setProperty('BordeDerCO'+estilo, borde_der.getColor().asRgbColor().asHexString());
        estilos_sheet.setProperty('BordeDerST'+estilo, borde_der.getBorderStyle());
      }
    

      }
    estilos_sheet.setProperty('colorLetra'+estilo, celda.getFontColor())
                 .setProperty('colorFondo'+estilo, celda.getBackground())
                 .setProperty('sizeFuente'+estilo, celda.getFontSize()+'');
    return {colorFondo: estilos_sheet.getProperty('colorFondo'+estilo),
            colorLetra: estilos_sheet.getProperty('colorLetra'+estilo),
            BordeSupCO: estilos_sheet.getProperty('BordeSupCO'+estilo),
            BordeSupST: estilos_sheet.getProperty('BordeSupST'+estilo),
            BordeInfCO: estilos_sheet.getProperty('BordeInfCO'+estilo),
            BordeInfST: estilos_sheet.getProperty('BordeInfST'+estilo),
            BordeIzqCO: estilos_sheet.getProperty('BordeIzqCO'+estilo),
            BordeIzqST: estilos_sheet.getProperty('BordeIzqST'+estilo),
            BordeDerCO: estilos_sheet.getProperty('BordeDerCO'+estilo),
            BordeDerST: estilos_sheet.getProperty('BordeDerST'+estilo)
            };

}

function cargarEstilos(){
  return estilos_sheet.getProperties();
}
function eliminarEstilo(estilo){

  estilos_sheet.deleteProperty('colorLetra'+estilo);
  estilos_sheet.deleteProperty('colorFondo'+estilo);
  estilos_sheet.deleteProperty('sizeFuente'+estilo);    


  estilos_sheet.deleteProperty('BordeSupCO'+estilo);
  estilos_sheet.deleteProperty('BordeSupST'+estilo);
  estilos_sheet.deleteProperty('BordeInfCO'+estilo);
  estilos_sheet.deleteProperty('BordeInfST'+estilo);
  estilos_sheet.deleteProperty('BordeIzqCO'+estilo);
  estilos_sheet.deleteProperty('BordeIzqST'+estilo);
  estilos_sheet.deleteProperty('BordeDerCO'+estilo);
  estilos_sheet.deleteProperty('BordeDerST'+estilo);
 
}
function borrarEstilos(){
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear({formatOnly:true});
}
function borrarTodo(){
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear();
}