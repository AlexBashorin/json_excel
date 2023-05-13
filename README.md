# json to excel & excel to json
Microservice accepting JSON as input and returning a file in xlsx format.  
And generate json from submitted.

`For install json to excel:`  
docker pull albaxeshtest1/json_xlsx:jexcel
docker build -t -d -p 3030:8082 albaxeshtest1/json_xlsx:jexcel

`For install excel to json:`  
docker pull albaxeshtest1/xlsx_json:excel_json  
docker run -d -it -p 4422:4444 albaxeshtest1/xlsx_json:excel_json  


