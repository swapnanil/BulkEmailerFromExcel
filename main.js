function validateEmail(email) {
    var re = /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(email);
}

var mandrill = require('mandrill-api/mandrill');
var mandrill_client = new mandrill.Mandrill('2t6q7YkVFC4KvKdlkC6sRg');

var message = {};
fs = require('fs')
fs.readFile('content.html', 'utf8', function (err, data) {
    message = {
        "text" : data,
        "html" : data,
        "subject": "the beatles",
        "from_email": "fly@confluenceedu.com",
        "from_name": "__PRO__",
        "to" : [],
        "headers": {
            "Reply-To": "fly@confluenceedu.com"
        }
    };
    readAndSend();
});

function readAndSend() {
    if(typeof require !== 'undefined') XLSX = require('xlsx');
    var workbook = XLSX.readFile('test.xlsx');
    var sheet_name_list = workbook.SheetNames;
    sheet_name_list.forEach(function(y) {
      var worksheet = workbook.Sheets[y];
      for (z in worksheet) {
        if(z[0] === '!') continue;
        if(validateEmail(worksheet[z].v)) {
	        message['to'].push({
	                "email": worksheet[z].v,
	                "name": "Cool guys and gals",
	                "type": "to"
	            });
    	}
        console.log("Email to: " + worksheet[z].v);
      }
    });
    mandrill_client.messages.send({"message": message}, function(result) {
	    console.log(result);
	}, function(e) {
	    console.log('A mandrill error occurred: ' + e.name + ' - ' + e.message);
	});
}
