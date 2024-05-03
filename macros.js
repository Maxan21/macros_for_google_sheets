/** @OnlyCurrentDoc */

function onEdit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getDataRange(); // Получаем диапазон ячеек с данными
  var values = range.getValues(); // Получаем значения ячеек
  
  var colors = ['#f4cccc', '#c9daf8', '#d9ead3', '#fff2cc', '#ead1dc', '#cfe2f3', '#d9d2e9', '#d0e0e3', '#fbe5d6', '#CD853F', '#1E90FF', '#DDA0DD', '#BDB76B', '#FFFF00', '#66CDAA', '#FF1493']; // Список цветов для выделения
  var k = 0; // Индекс для цветов
  var proxy = []; // Массив для хранения уже обработанных значений
  // Проходим по всем ячейкам
  for (var row = 1; row < values.length; row++) {
      for (var col = 5; col < values[row].length; col += 6) {
          var value = values[row][col];
          if (value && (typeof value === 'string' || typeof value === 'number') && !proxy.includes(value)) {
              var flag = 0;
              var str=[row+1];
              var colomn=[col+1] // Добавляем первоначальную ячейку для выделения
              for (var rw = 1; rw < values.length; rw++) {
                  for (var cl = 5; cl < values[rw].length; cl += 6) {
                      var value1 = values[rw][cl];
                      if (rw !== row || cl !== col) { // Исключаем первоначальную ячейку из сравнения
                          if (value === value1) {
                              str.push(rw+1);
                              colomn.push(cl+1) // Добавляем ячейку дубликат
                              if  (values[row][col - 3] !== values[rw][cl - 3]){
                                flag=1;
                              }
                          }
                      }
                  }
              }
              
              if (flag===1) {
                  for (var i=0; i < str.length; i++){
                    sheet.getRange(str[i], colomn[i]).setBackground(colors[k]);
                  }
                  proxy.push(value); // Добавляем значение в массив уже обработанных
                  k++; // Переходим к след
              }
                            
              else {
                  for (var i=0; i < str.length; i++){
                    sheet.getRange(str[i], colomn[i]).setBackground('#fff');
                  }
                  proxy.push(value);
                   // Добавляем значение в массив уже обработанных
                   // Переходим к след
              }

            }
          }
  }
}