const applyHeaderStyle = async (sheet, cell, value) => {
    sheet.cell(cell)
      .value(value)
      .style({
       fontFamily: "Toyota Type Light",           
        fontSize: 14,                  
        bold: true,                    
        italic: false,                  
        underline: false,                      
        fill: {
          type: "solid",               
          color: "FFFF00"              
        },
        horizontalAlignment: "center", 
        verticalAlignment: "middle",   
        border: true                  
      });
  };
  const applyCellStyle = async (sheet, cell, value,verticalCenter=false) => {
    sheet.cell(cell)
      .value(value)
      .style({
        fontFamily: "Toyota Type Light",          
        fontSize: 12,                  
        bold: false,                    
        italic: false,                  
        underline: false,                 
        border: true                
      });
  };

  // Method to set row height and column width
 const setRowHeightAndColumnWidth = async (sheet, row, height, column, width) => {
    sheet.row(row).height(height);
    sheet.column(column).width(width);
  }


module.exports ={setRowHeightAndColumnWidth,applyHeaderStyle,applyCellStyle}
