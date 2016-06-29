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
    var datalength = 0;
    // Ready The Event For When A File Gets Selected
    reader.onload = function(e) {
        $('.overlay').show();
        var data = e.target.result;
        var cfb = XLS.CFB.read(data, {type: 'binary'});
        var wb = XLS.parse_xlscfb(cfb);
        // Loop Over Each Sheet
        
            // Obtain The Current Row As CSV
            var sCSV = XLS.utils.make_csv(wb.Sheets['Sheet1']);   
            var data = XLS.utils.sheet_to_json(wb.Sheets['Sheet1'], {header:1});
            console.log(data);
            var newdata = [];
            
            data.forEach(function (arr) {
                if (arr.length > 0) {
                    newdata.push(arr);
                }
            });
            
            var sampleObject = {};
            
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
                    deliveryDate: '',
                    firstName: ''
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
                cData.firstName = col.CUSTOMER_NAME;
                cData.deliveryDate = col.CLOSEDATE;

                productname = ''

                stringifyCData = JSON.stringify(cData);

                objectToPost.push({
                    emailId: col.CUSTOMER_EMAIL,
                    firstName: col.CUSTOMER_NAME,
                    lastName: '',
                    cData: stringifyCData
                })
            });
            var datalength = objectToPost.length;
            var index = 0;
            
            postRequest(index);
            
            function postRequest (index) {
                if (index < datalength) {
                    var data = objectToPost[index];
                    $.post('http://jabong.apitest.zykrr.com/token/12', data).done(function (result) {
                        objectToPost[index].token = result.uid; 
                        objectToPost[index].url = "http://jabong.zykrr.com?token=" + result.uid;
                        index++;
                        var progress = index / datalength * 100;
                        $('.pval').html(progress + '% ');
                        $('.progress-bar').css('width', progress+'%').attr('aria-valuenow', progress);
                        postRequest(index);
                    })
                } else {
                    setNewData(index, datalength);
                    $('.overlay').hide();
                }
            }

            function setNewData (i, dl) {
                var oIndex = 0;
                objectToPost.forEach(function (postData) {
                    postData.cData = JSON.parse(postData.cData);
                    newdata[oIndex].push(postData.token);
                    newdata[oIndex].push(postData.url);
                    oIndex++;
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
                    if (i === dl) {
                        $("#my_file_output").append(sRow);    
                    }
                });
            }
    };
    
    // Tell JS To Start Reading The File.. You could delay this if desired
    reader.readAsBinaryString(oFile);
}

});//]]> 