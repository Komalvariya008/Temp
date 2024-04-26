const qr = require("qrcode");
let data = {
"name":"komal variya",
"email": "variyakomal008@gmail.com",
"gender": "female",
"id":123
};
let stJson = JSON.stringify(data);
qr.toString(stJson, {type:"terminal"},function(err, code)
{
if(err)  return console.log("error");
console.log(code);
});
// qr.toDataURL(stJson,function(err, code)
// {
// if(err)  return console.log("error");
// console.log(code);
// });