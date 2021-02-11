/**Made By Lucas Fenstra (Hogo)
     * 
     * THIS IS FROM MY GIST, FOR THESE STEPS YOU GOT TO GO TO MY GIST:https://gist.github.com/hogo1510/a771ca5020bce803bf9c99baaf0cd1ef
     * OR JUST COPY THE CODE AND PUT IT IN YOUR SCRIPT LAB>
     * 
     * It's made in Script Lab.
     * if you want to use this then do this >
     * you can copy the link of the gist, then go to excel, then go to Script Lab (If you don't Have Script Lab Them See This: https://appsource.microsoft.com/en-us/product/office/wa104380862)
     * press on Import, and put the link of the gist there.
     * then you need to press the run option (from the script lab tab)
     * and the put the password: "Test!" in and the press "confirm" and then "Run"
     * That's it
     * 
     * you can't open it as a YAML File because you will need te run option of Script Lab, althoug i'm trying to fix it so you don't need script lab.
     * but for now you will need Script Lab.
     * If you found a way that you don't need Script Lab, then please let me know.
    */
   $("#run").click(() => tryCatch(run));
   //sets the password checker on false so you can't press run
   var checkPass = false;
   function btn() {
     var passwText;
     var passwrd = document.getElementById("InputPassw").value;
     switch (passwrd) {
       //Password Here!
       case "Test!": // <---
         //what happends if the passw is Correct
         passwText = "Correct";
         checkPass = true;
         break;
       //what happpends if the passw is incorect
       default:
         passwText = "Incorect!";
         checkPass = false;
     }
     document.getElementById("Message").innerHTML = passwText;
   }
   async function run() {
     await Excel.run(async (context) => {
       const sheet = context.workbook.worksheets.getActiveWorksheet();
       /**Code For Excel Here! */
       //checks if the password is correct
       if (checkPass == true) {
         //the code if the passw is correct
         console.log("Nice That's the good password, u r a Legend");
         console.log("If u see This Then It's working, btw the code sucks tho");
         
         //Headers
         sheet.tables.add("B2:E5", true);
         var headers = [["Product", "hoeveelheid", "Prijs Per Unit", "Totalen"]];
         var headerRange = sheet.getRange("B2:E2");
         headerRange.values = headers;
         headerRange.format.fill.color = "#4472C4";
         headerRange.format.font.color = "white";
         
         //Products
         var productData = [["Noten", 6, 7.5], ["koffie", 20, 34.5], ["Chocolade", 10, 9.56]];
         var dataRange = sheet.getRange("B3:D5");
         dataRange.values = productData;
         //Formu Amount Sold (aka Basic Math lol)
         var totalFormulas = [["=C3 * D3"], ["=C4 * D4"], ["=C5 * D5"], ["=SUM(E3:E5)"]];
         var totatRange = sheet.getRange("E3:E6");
         totatRange.formulas = totalFormulas;
         totatRange.format.font.bold = true;
         //Total in US Dollars
         totatRange.numberFormat = [["$0.00"]];
         /**The End, Lmao */
         await context.sync();
       } else {
         //the code if u have the wrong passw
         console.log("Sike That's the wrong password, lol");
       }
     });
   }
   //Just Don't touch this :)
   /** Default helper for invoking an action and handling errors. */
   async function tryCatch(callback) {
     try {
       await callback();
     } catch (error) {
       // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
       console.error(error);
     }
   }