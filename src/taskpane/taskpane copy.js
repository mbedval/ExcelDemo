/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */
import fetch from 'node-fetch';
//import fetch from 'cross-fetch';

Office.onReady(info => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    await Excel.run(async context => {
      
      //
       const Authorization = "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiIsIng1dCI6InBpVmxsb1FEU01LeGgxbTJ5Z3FHU1ZkZ0ZwQSIsImtpZCI6InBpVmxsb1FEU01LeGgxbTJ5Z3FHU1ZkZ0ZwQSJ9.eyJhdWQiOiI3MTdiNzc4Ny05NTdhLTQwZWUtYTk5ZS1iMjIyNDAwNWQwODgiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC83YjQwNjdhMi1iZjFmLTQ0MjgtODkxZi00M2E3ZDBjZTdjYTkvIiwiaWF0IjoxNTc4NTYyNjk0LCJuYmYiOjE1Nzg1NjI2OTQsImV4cCI6MTU3ODU2NjU5NCwiYWlvIjoiNDJOZ1lMaldiRGozNVgvdTcvK3ZMVEc2WlhHU0J3QT0iLCJhcHBpZCI6IjIxNzc5ZDI2LTFjYmEtNDQzMC1hMjRmLTNmYjc3YTJlODNmMyIsImFwcGlkYWNyIjoiMSIsImlkcCI6Imh0dHBzOi8vc3RzLndpbmRvd3MubmV0LzdiNDA2N2EyLWJmMWYtNDQyOC04OTFmLTQzYTdkMGNlN2NhOS8iLCJvaWQiOiIyYzgyMzA4Ni04ZTAyLTRmNzktYTdiYy0wMjcwYzk3YWE2YzkiLCJzdWIiOiIyYzgyMzA4Ni04ZTAyLTRmNzktYTdiYy0wMjcwYzk3YWE2YzkiLCJ0aWQiOiI3YjQwNjdhMi1iZjFmLTQ0MjgtODkxZi00M2E3ZDBjZTdjYTkiLCJ1dGkiOiJRYzFtREFOdXdrMlpzbVlCWk9XaUFBIiwidmVyIjoiMS4wIn0.x9a8RVNWsq7Z2JomY3SOyPJFV6rxnjJ-ATzHs-XJGn8aGks6CgFIY1vkOw2I-4Ii7bOuIyhzqRoFXNjtdJskct0HdE1qkmaKzHmTZelKiToudsRE5YQLc80gtFRfLfW5kJN_K7B7d3EM2V-_i4YLOFYaRMXKzlQuXIySS8TmhRJEPpqgco6rkV6oGWKTGDF_TjzezI_D8dDJEwqsLBpBU2tf_re4x7NDgx0s2pVy8jLMpsKp3w7DlwADlgDZubksybLqnhCcQvY6vk3bpUFgfqnCL8qPLPWmQ0EqHvMLmcWFOtmU1aopMoaotE_ZLOZWYKk1C5zPrUNduZj6Dtg7Kg";
       const OcpApimSubscriptionKey = "c732351ad851400aa9fd46a76cf843ab";
       const url = "https://use2-impact-cdpgateway.azure-api.net/api/CheckList";
        //const data = { name: "Mukesh123" };
      //

      fetch(url,{
        method: 'GET',
        mode: 'cors',
        headers: {  
          'Authorization': Authorization,
          'Ocp-Apim-Subscription-Key': OcpApimSubscriptionKey
        }
      }
        )
      .then(res => {   
        if (res.status == 401) {
          throw new Error("Authorization has been denied for this request.");
        }        
        return res.json();
      })
      .then(user => {
        console.log(user);
      })
      .catch(err => {
 
        console.error(err);
      });

      //const range = context.workbook.getSelectedRange();

      var activeCell = context.workbook.worksheets.getItem("Sheet1").getRange("A1:B2");

      activeCell.load("values");
      await context.sync().then(function(){      
        console.log(activeCell.values[0][1])
      });

    

    
    });
  } catch (error) {
    console.error(error);
  }
}
