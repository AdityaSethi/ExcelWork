$(window).load(function(){
var oFileIn;

$(function() {
    oFileIn = document.getElementById('my_file_input');
    if(oFileIn.addEventListener) {
        oFileIn.addEventListener('change', filePicked, false);
    }

    $('#export_button').bind('click', function () {
        $("#my_file_output").table2excel({
            exclude: ".noExl",
            name: "Excel Document Name",
            filename: "Jabong Data",
            fileext: ".xls",
            exclude_img: true,
            exclude_links: true,
            exclude_inputs: true
        });
    })
});

function filePicked(oEvent) {
    // Get The File From The Input
    var oFile = oEvent.target.files[0];
    var sFilename = oFile.name;
    // Create A File Reader HTML5
    var reader = new FileReader();
    
    // Ready The Event For When A File Gets Selected
    reader.onload = function(e) {
        var data = e.target.result;
        var cfb = XLS.CFB.read(data, {type: 'binary'});
        var wb = XLS.parse_xlscfb(cfb);
        // Loop Over Each Sheet
        wb.SheetNames.forEach(function(sheetName) {
            // Obtain The Current Row As CSV
            var sCSV = XLS.utils.make_csv(wb.Sheets[sheetName]);   
            var data = XLS.utils.sheet_to_json(wb.Sheets[sheetName], {header:1});
            
            var newdata = [];
            
            data.forEach(function (arr) {
                if (arr.length > 0) {
                    newdata.push(arr);
                }
            });
            
            var sampleObject = {};
            console.log(newdata);
            newdata[0].forEach(function (key) {
                sampleObject[key] = ''
            });

            var header = newdata.shift();
            var collection = [];
            
            newdata.forEach(function (nd) {
                var sampleObjectTemp = JSON.parse(JSON.stringify(sampleObject));
                var i = 0;
                for (key in sampleObjectTemp) {
                    sampleObjectTemp[key] = nd[i];
                    i++
                }
                i = 0;
                collection.push(sampleObjectTemp);
            })

            var objectToPost = [];
            collection.forEach(function (col) {
                
                var cData = {
                    productName: '',
                    orderNo: '',
                    productImg: 'todo',
                    orderDate: '',
                    deliveryDate: ''
                };
                var productname = '';
                for (var key in col) {
                    if (key.indexOf('SKU') !== -1) {
                        if (col[key]) {
                            productname = col[key] + ',' + productname;
                        }
                    }
                }
                cData.productName = productname.substring(0, productname.length - 1);
                cData.orderNo = col.ORDERNO;
                cData.orderDate = col.ORDERDATE;

                productname = ''

                stringifyCData = JSON.stringify(cData);

                objectToPost.push({
                    emailId: col.CUSTOMER_EMAIL,
                    firstName: col.CUSTOMER_NAME,
                    lastName: '',
                    cData: stringifyCData,
                })
            });
            var datalength = objectToPost.length;
            objectToPost.forEach(function (postObj) {
                $.post('http://jabong.apitest.zykrr.com/token/12', postObj).done(function (data) {
                    postObj.token = data.uid;
                    postObj.url = "http://jabong.zykrr.com?token=" + data.uid + '%' + data.emailId;
                    datalength--;
                    if (datalength == 0) {
                        objectToPost.forEach(function (postData) {
                            postData.cData = JSON.parse(postData.cData);
                            newdata.forEach(function (ndata) {
                                if (ndata[0] === postData.cData.orderNo) {
                                    ndata.push(postData.token);
                                    ndata.push(postData.url);
                                }
                            })
                        });

                        header.push('TOKEN');
                        header.push('URL');

                        newdata.unshift(header);

                        $.each(newdata, function( indexR, valueR ) {
                            var sRow = "<tr>";
                            $.each(newdata[indexR], function( indexC, valueC ) {
                                sRow = sRow + "<td>" + valueC + "</td>";
                            });
                            sRow = sRow + "</tr>";
                            $("#my_file_output").append(sRow);
                        });
                    }
                })
            });
        });
    };
    
    // Tell JS To Start Reading The File.. You could delay this if desired
    reader.readAsBinaryString(oFile);
}

});//]]> 