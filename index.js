const express = require('express')
const bodyParser = require('body-parser')
const knex = require('knex')(require('./knexfile'))
const crypto = require('crypto')
var xl = require('excel4node')

const app = express()
app.use(express.static(__dirname + '/public'))
app.use(bodyParser.json())

var path = require('path')

  app.get('/proceso', function(req, res) {
    var wb = new xl.Workbook();
 
    // Add Worksheets to the workbook
    var ws = wb.addWorksheet('Sheet', {
      view: {
        zoom: 200
      }
    });
    var wsGraph = wb.addWorksheet('Result Graph');

    ws.row(1).freeze();   
    // Create a reusable style
    
    var style = wb.createStyle({
      font: {
        color: '#000000',
        size: 12
      },
      alignment: {
        vertical: 'center',
        horizontal: 'center'
      }
    });

    var percentStyle = wb.createStyle({
      font: {
        color: '#000000',
        size: 12,
      },
      alignment: {
        vertical: 'center',
        horizontal: 'center'
      },
      numberFormat: '#.00%; -#.00%; -'
    });
    ws.row(1).setHeight(150);
    for( var i = 1; i < 10; i ++ )
      ws.column(i).setWidth(15);
    ws.column(i).setWidth(35);
    ws.cell(1, 1).string("Folio").style(style);
    ws.cell(1, 2).string("SEXO (Hombre=1 / Mujer=2)").style(style);
    ws.cell(1, 3).string("1. Fuma\n Cigarrillos (0=no\n fumo\n 1=0-5 a la\n semana 2=0 a 5\n al día 3=6 a 10 al\n día 4=11 a 20 al\n día 5=más de 1\n cajetilla al día)").style(style);
    ws.cell(1, 4).string("smoke").style(style);
    ws.cell(1, 5).string("2. Moderate Physical Activity 30 Mnts in free time(yes=1 / No=0)").style(style);
    ws.cell(1, 6).string("Act física").style(style);
    ws.cell(1, 7).string("3. Consumo mayor a 5 porciones de Frutas y verduras diarias (SI=1 / No=0)").style(style);
    ws.cell(1, 8).string("Frutas y verduras").style(style);
    ws.cell(1, 9).string("4. Frecuencia Bebida Alcoholica (0=nunca 1=una o menos veces al mes 2=2 a 4 veces por mes 3=2 a 3 veces por semana 4=4 o más veces por semana)").style(style);
    ws.cell(1, 10).string("5. Cantidad de tragos en un día normal (0= 1 o 2 1= 3 o 4 2=5 o 6 3=7, 8 o 9 4=10 o más 0= no tomo bebidas alcohólicas)").style(style);
    ws.cell(1, 11).string("6. Frecuencia de tragos en un solo día (0=nunca 1=menos de una vez  al mes 2=mensualmente 3=semanalmente 4=todos los días o casi todos los días 0=no tomo bebidas alcohólicas)").style(style);
/* ------------------------------- Table One -------------------------------- */
    wsGraph.column(2).setWidth(40);
    wsGraph.column(7).setWidth(40);
    wsGraph.row(5).setHeight(30);

    wsGraph.cell(3, 3).string("Frecuencias").style(style);
    wsGraph.cell(3, 8).string("Porcentajes").style(style);
    
    wsGraph.cell(5, 2).string("Pregunta 1. ¿Actualmente usted fuma cigarrillos?").style(style);
    wsGraph.cell(5, 3).string("Hombre").style(style);
    wsGraph.cell(5, 4).string("Mujer").style(style);
    wsGraph.cell(5, 5).string("Total").style(style);

    wsGraph.cell(5, 7).string("Pregunta 1. ¿Actualmente usted fuma cigarrillos?").style(style);
    wsGraph.cell(5, 8).string("Hombre").style(style);
    wsGraph.cell(5, 9).string("Mujer").style(style);
    wsGraph.cell(5, 10).string("Total").style(style);

    wsGraph.cell(6, 2).string("1=0-5 a la semana").style(style);
    wsGraph.cell(7, 2).string("2=0 a 5 al día").style(style);
    wsGraph.cell(8, 2).string("3=6 a 10 al día").style(style);
    wsGraph.cell(9, 2).string("4=11 a 20 al día").style(style);
    wsGraph.cell(10, 2).string("5=más de 1 cajetilla al día)").style(style);
    wsGraph.cell(11, 2).string("SI").style(style);
    wsGraph.cell(12, 2).string("NO").style(style);

    wsGraph.cell(6, 7).string("1=0-5 a la semana").style(style);
    wsGraph.cell(7, 7).string("2=0 a 5 al día").style(style);
    wsGraph.cell(8, 7).string("3=6 a 10 al día").style(style);
    wsGraph.cell(9, 7).string("4=11 a 20 al día").style(style);
    wsGraph.cell(10, 7).string("5=más de 1 cajetilla al día)").style(style);
    wsGraph.cell(11, 7).string("SI").style(style);
    wsGraph.cell(12, 7).string("NO").style(style);

    wsGraph.cell(13, 2).string("Total general").style(style);
    wsGraph.cell(13, 7).string("Total general").style(style);
/* ------------------------------- Table Two -------------------------------- */
    wsGraph.row(17).setHeight(150);

    wsGraph.cell(17, 2).string("Pregunta 2. ¿En el último mes practicó deporte o realizó actividad física fuera de su horario de trabajo durante 30 minutos o más de forma regular (al menos tres veces por semana)?").style(style);
    wsGraph.cell(17, 3).string("Hombre").style(style);
    wsGraph.cell(17, 4).string("Mujer").style(style);
    wsGraph.cell(17, 5).string("Total").style(style);

    wsGraph.cell(18, 2).string("SI").style(style);
    wsGraph.cell(19, 2).string("NO").style(style);

    wsGraph.cell(17, 7).string("Pregunta 2. ¿En el último mes practicó deporte o realizó actividad física fuera de su horario de trabajo durante 30 minutos o más de forma regular (al menos tres veces por semana)?").style(style);
    wsGraph.cell(17, 8).string("Hombre").style(style);
    wsGraph.cell(17, 9).string("Mujer").style(style);
    wsGraph.cell(17, 10).string("Total").style(style);

    wsGraph.cell(18, 7).string("SI").style(style);
    wsGraph.cell(19, 7).string("NO").style(style);

    wsGraph.cell(20, 2).string("Total general").style(style);
    wsGraph.cell(20, 7).string("Total general").style(style);
/* ------------------------------- Table Three -------------------------------- */

    wsGraph.row(24).setHeight(100);

    wsGraph.cell(24, 2).string("Pregunta 3. Si toma en cuenta todas las verduras y frutas que come en el día, ¿Suman 5 porciones?").style(style);
    wsGraph.cell(24, 3).string("Hombre").style(style);
    wsGraph.cell(24, 4).string("Mujer").style(style);
    wsGraph.cell(24, 5).string("Total").style(style);

    wsGraph.cell(25, 2).string("SI").style(style);
    wsGraph.cell(26, 2).string("NO").style(style);

    wsGraph.cell(24, 7).string("Pregunta 3. Si toma en cuenta todas las verduras y frutas que come en el día, ¿Suman 5 porciones?").style(style);
    wsGraph.cell(24, 8).string("Hombre").style(style);
    wsGraph.cell(24, 9).string("Mujer").style(style);
    wsGraph.cell(24, 10).string("Total").style(style);

    wsGraph.cell(25, 7).string("SI").style(style);
    wsGraph.cell(26, 7).string("NO").style(style);

    wsGraph.cell(27, 2).string("Total general").style(style);
    wsGraph.cell(27, 7).string("Total general").style(style);

    var result = '';
    var total = 0;
    /* --------------------------------- smoke init ----------------------------- */
    var smoke = new Array(2)
    for( var j = 0; j < 2; j++ )
    {
      smoke[j] = []
      for( var k = 0; k < 7; k++ )
        smoke[j][k] = 0;
    }
    /* --------------------------------- sport init ----------------------------- */
    var sport = new Array(3).fill(0).map(() => new Array(3).fill(0));
    knex.select('*')
        .from('result_9f')
        .then((rows) => {
          var i = 2;
          for (row of rows) {
            index = row['id'];
            var gender = row['d1'] - 1
            /* --------------------------------------------------------- worksheet 1 --------------------------------------------------------- */
            ws.cell(i, 1).number(row['id']).style(style);
            ws.cell(i, 2).number(row['d1']).style(style);

            if( row['p1'] == null ) {
              ws.cell(i, 3).string('').style(style);
              smoke[gender][0] = smoke[gender][0] + 1;
            }
            else if( row['p1'] == '0' )
            {
              ws.cell(i, 3).string('NO').style(style);
              smoke[gender][0] = smoke[gender][0] + 1;
            }
            else
            {
              ws.cell(i, 3).number(row['p1']).style(style);
              smoke[gender][row['p1']] = smoke[gender][row['p1']] + 1;
              smoke[gender][6] += 1;
            }

            if( row['p1'] == null || row['p1'] == 0)
              ws.cell(i, 4).string('NO').style(style);
            else
              ws.cell(i, 4).string('SI').style(style);

            ws.cell(i, 5).string('').style(style);

            if( row['p2'] == '0' )
            {
              sport[1][gender]++; sport[1][2]++; sport[2][gender]++;
              ws.cell(i, 6).string('NO').style(style);
            }
            else
            {
              sport[0][gender]++; sport[0][2]++; sport[2][gender]++;
              ws.cell(i, 6).string('SI').style(style);
            }

            if( row['p3'] == null )
              ws.cell(i, 7).string('').style(style);
            else if( row['p3'] == '0' )
              ws.cell(i, 7).string('NO').style(style);
            else
              ws.cell(i, 7).number(row['p3']).style(style);

            if( row['p4'] == null )
              ws.cell(i, 9).string('').style(style);
            else
              ws.cell(i, 9).number(row['p4']).style(style);

            if( row['p5'] == null )
              ws.cell(i, 10).string('').style(style);
            else
              ws.cell(i, 10).number(row['p5']).style(style);
            
            if( row['p6'] == null )
              ws.cell(i, 11).string('').style(style);
            else
              ws.cell(i, 11).number(row['p6']).style(style);


            i++;
          }
/* ------------------------------ Frequency one ----------------------------------- */
          for( var k = 1; k <= 6; k ++ )
          {
            wsGraph.cell(k + 5, 3).number(smoke[0][k]).style(style);
            wsGraph.cell(k + 5, 4).number(smoke[1][k]).style(style);
            wsGraph.cell(k + 5, 5).number(smoke[0][k] + smoke[1][k]).style(style);
          }

          wsGraph.cell(12, 3).number(smoke[0][0]).style(style);
          wsGraph.cell(12, 4).number(smoke[1][0]).style(style);
          wsGraph.cell(12, 5).number(smoke[0][0] + smoke[1][0]).style(style);
          wsGraph.cell(13, 3).formula('C11+C12').style(style);
          wsGraph.cell(13, 4).formula('D11+D12').style(style);
          wsGraph.cell(13, 5).formula('C13+D13').style(style);

/* ------------------------------ Percentage one ----------------------------------- */

          total = smoke[0][6] + smoke[1][6] + smoke[0][0] + smoke[1][0];
          for( var k = 1; k <= 6; k ++ )
          {
            wsGraph.cell(k + 5, 8).number(smoke[0][k] / total).style(percentStyle);
            wsGraph.cell(k + 5, 9).number(smoke[1][k] / total).style(percentStyle);
            wsGraph.cell(k + 5, 10).number((smoke[0][k] + smoke[1][k]) / total).style(percentStyle);
          }

          wsGraph.cell(12, 8).number(smoke[0][0] / total).style(percentStyle);
          wsGraph.cell(12, 9).number(smoke[1][0] / total).style(percentStyle);
          wsGraph.cell(12, 10).number((smoke[0][0] + smoke[1][0]) / total).style(percentStyle);
          wsGraph.cell(13, 8).formula('H11+H12').style(percentStyle);
          wsGraph.cell(13, 9).formula('I11+I12').style(percentStyle);
          wsGraph.cell(13, 10).formula('J11+J12').style(percentStyle);
/* ------------------------------ Frequency Two ----------------------------------- */
          for(j = 0; j < 3; j++)
            for(k = 0; k < 3; k++)
              wsGraph.cell(18 + j, 3 + k).number(sport[j][k]).style(style);
          wsGraph.cell(20, 5).number(sport[0][2] + sport[1][2]).style(style);

/* ------------------------------ Percentage two ----------------------------------- */
          for(j = 0; j < 3; j++)
            for(k = 0; k < 3; k++)
              wsGraph.cell(18 + j, 8 + k).number(sport[j][k] / total).style(percentStyle);
          wsGraph.cell(20, 10).number((sport[0][2] + sport[1][2]) / total).style(percentStyle);

          wb.write('Excel1.xlsx'); 
        }
      )
    /* --------------------------------------------------------- worksheet 2 --------------------------------------------------------- */

    

    res.send("Created");
  });
//______________ port listen ________________________

app.listen(7555, () => {
  console.log('Server running on http://localhost:7555')
})