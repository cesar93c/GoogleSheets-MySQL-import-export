

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('MySQL')
      .addSubMenu(ui.createMenu('Exportar datos a la tabla')
          .addItem('Enviar y Agregar a tabla', 'sqlAgregar')
          .addItem('Enviar y Reemplazar tabla', 'sqlReemplazar'))
      .addSeparator()
      .addItem('Importar datos de la tabla', 'MySQLFetchData')
      .addToUi();
}

//IMPORTANTE: ANTES DEBE ESTAR CREADA LA TABLA EN LA BASE. SI PREFIERE, PUEDE CREARSE EN EL SCRIPT "CREAR TABLA", O SINO DIRECTAMENTE EN PHPMYADMIN

function sqlAgregar() {
  //Get spreadsheet and spreadsheet row length

  var activeSs = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet = SpreadsheetApp.getActiveSheet()
  var numSsRows = Sheet.getLastRow(); //Store last row to get define range for export
  var columnsRange = Sheet.getLastColumn();

  //cache values of status data range for export
  var range = Sheet.getRange(2, 1, numSsRows, columnsRange)
     .getValues();


  //-----------PONER DATOS DE CONEXION---------------------------------------------------------------------------------
  // jdbc connect variables. CREAR UNA HOJA QUE SE LLAME conn Y PONER LOS DATOS DE CONEXION EN LAS CELDAS MENCIONADAS ABAJO
  var address = activeSs.getSheetByName('conn').getRange('C2').getDisplayValue();
  var user = activeSs.getSheetByName('conn').getRange('C3').getDisplayValue();
  var userPwd = activeSs.getSheetByName('conn').getRange('C4').getDisplayValue();
  var db = activeSs.getSheetByName('conn').getRange('C5').getDisplayValue();
  var table = activeSs.getSheetByName('conn').getRange('C6').getDisplayValue();


  //Connect to remote db
  var dbUrl = 'jdbc:mysql://' + address + ':3306/' + db;


    var conn = Jdbc.getConnection(dbUrl, user, userPwd);
    conn.setAutoCommit(false);

  //-----------PONER LOS HEADERS QUE CORRESPONDAN A LA TABLA---------------------------------------------------------------
  //Commit all rows to MySQL database  -
  var stmt = conn.prepareStatement('INSERT INTO '+table+' ' + '(column_1, column_2, column_3, column_4, column_5) values (?, ?, ?, ?, ?)'); //SQL statement to write batch to db
  for (var i = 0; i < (numSsRows-1); i++) {
    stmt.setString(1, range[i][0]); //Column_1
    stmt.setString(2, range[i][1]); //Column_2
    stmt.setString(3, range[i][2]); //Column_3
    stmt.setString(4, range[i][3]); //Column_4
    stmt.setString(5, range[i][4]); //Column_5
    stmt.addBatch();
  }

  stmt.executeBatch();
  conn.commit();
  conn.close();

}



//IMPORTANTE: ANTES DEBE ESTAR CREADA LA TABLA EN LA BASE. SI PREFIERE, PUEDE CREARSE EN EL SCRIPT "CREAR TABLA", O SINO DIRECTAMENTE EN PHPMYADMIN

function sqlReemplazar() {
  //Get spreadsheet and spreadsheet row length

  var activeSs = SpreadsheetApp.getActiveSpreadsheet();
  var Sheet = SpreadsheetApp.getActiveSheet()
  var numSsRows = Sheet.getLastRow(); //Store last row to get define range for export
  var columnsRange = Sheet.getLastColumn();

  //cache values of status data range for export
  var range = Sheet.getRange(2, 1, numSsRows, columnsRange)
     .getValues();


  //-----------PONER DATOS DE CONEXION---------------------------------------------------------------------------------
  // jdbc connect variables
  var address = activeSs.getSheetByName('conn').getRange('C2').getDisplayValue();
  var user = activeSs.getSheetByName('conn').getRange('C3').getDisplayValue();
  var userPwd = activeSs.getSheetByName('conn').getRange('C4').getDisplayValue();
  var db = activeSs.getSheetByName('conn').getRange('C5').getDisplayValue();
  var table = activeSs.getSheetByName('conn').getRange('C6').getDisplayValue();



  //Connect to remote db
  var dbUrl = 'jdbc:mysql://' + address + ':3306/' + db;


    var conn = Jdbc.getConnection(dbUrl, user, userPwd);
    conn.setAutoCommit(false);

  //Clear table data prior to importing latest
  var clearTable = conn.prepareStatement('TRUNCATE '+table+';');
  clearTable.addBatch();
  clearTable.executeBatch();



  //-----------PONER LOS HEADERS QUE CORRESPONDAN A LA TABLA---------------------------------------------------------------
  //Commit all rows to MySQL database  -
  var stmt = conn.prepareStatement('INSERT INTO '+table+' ' + '(column_1, column_2, column_3, column_4, column_5) values (?, ?, ?, ?, ?)'); //SQL statement to write batch to db
  for (var i = 0; i < (numSsRows-1); i++) {
    stmt.setString(1, range[i][0]); //Column_1
    stmt.setString(2, range[i][1]); //Column_2
    stmt.setString(3, range[i][2]); //Column_3
    stmt.setString(4, range[i][3]); //Column_4
    stmt.setString(5, range[i][4]); //Column_5
    stmt.addBatch();
  }

  stmt.executeBatch();
  conn.commit();
  conn.close();

}



// MySQL to Google Spreadsheet By Pradeep Bheron
// Support and contact at pradeepbheron.com

function MySQLFetchData() {

  var activeSs = SpreadsheetApp.getActiveSpreadsheet();
  var address = activeSs.getSheetByName('conn').getRange('C2').getDisplayValue();
  var user = activeSs.getSheetByName('conn').getRange('C3').getDisplayValue();
  var userPwd = activeSs.getSheetByName('conn').getRange('C4').getDisplayValue();
  var db = activeSs.getSheetByName('conn').getRange('C5').getDisplayValue();
  var table = activeSs.getSheetByName('conn').getRange('C6').getDisplayValue();



  var dbUrl = 'jdbc:mysql://' + address + ':3306/' + db;
  var conn = Jdbc.getConnection(dbUrl, user, userPwd);
    conn.setAutoCommit(false);

  var stmt = conn.createStatement();
  var start = new Date(); // Get script starting time

  var rs = stmt.executeQuery('SELECT * FROM '+table); // It sets the limit of the maximum nuber of rows in a ResultSet object

  //change table name as per your database structure

  var doc = SpreadsheetApp.getActiveSpreadsheet(); // Returns the currently active spreadsheet
  var cell = doc.getRange('a1');
  var row = 0;
  var getCount = rs.getMetaData().getColumnCount(); // Mysql table column name count.

  for (var i = 0; i < getCount; i++){
     cell.offset(row, i).setValue(rs.getMetaData().getColumnName(i+1)); // Mysql table column name will be fetch and added in spreadsheet.
  }

  var row = 1;
  while (rs.next()) {
    for (var col = 0; col < rs.getMetaData().getColumnCount(); col++) {
      cell.offset(row, col).setValue(rs.getString(col + 1)); // Mysql table column data will be fetch and added in spreadsheet.
    }
    row++;
  }

  rs.close();
  stmt.close();
  conn.close();
  var end = new Date(); // Get script ending time
  Logger.log('Time elapsed: ' + (end.getTime() - start.getTime())); // To generate script log. To view log click on View -> Logs.
}
