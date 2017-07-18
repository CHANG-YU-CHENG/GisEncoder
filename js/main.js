  var openFile = function(event) {
    var input = event.target;
    var reader = new FileReader();
    reader.onload = function(){
      var arrayBuffer = reader.result;
      var arr = fixdata(arrayBuffer); 
      output = []; 
      output.push(["location","longitude","latitude"]); 
      // import xlsx, xlsx to json: https://goo.gl/naHt6u
      var workbook = XLSX.read(btoa(arr), {type : 'base64'});
      workbook.SheetNames.forEach(function(sheetName){
        var XL_row_object = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
        for(var i=0;i< XL_row_object.length ;i++){
          var address = XL_row_object[i]['location']; 
          // gooogle map api: https://goo.gl/p7eNJi
          var geocoder = new google.maps.Geocoder();
          if (geocoder) {
            geocoder.geocode({ 'address': address }, function (results, status) {
               if (status == google.maps.GeocoderStatus.OK) {
                  var addr = results[0].formatted_address;
                  var lng = results[0].geometry.location.lng();
                  var lat = results[0].geometry.location.lat();
                  output.push([addr,lng,lat]);
               }
               else {
                  console.log("Geocoding failed: " + status);
               }
            });
          }                        
        }
      })
    };
    reader.readAsArrayBuffer(input.files[0]);
  };

  //global variable
  var output = []; 
  // export xlsx: https://goo.gl/YNxVZm
  var exportXlsx = function(){
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.aoa_to_sheet(output, {cellDates:true});
    XLSX.utils.book_append_sheet(wb, ws, "SheetJS");
    var wbout = XLSX.write(wb, {bookType: 'xlsx' , bookSST: true, type: 'binary'});
    saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), "GisEncodedData.xlsx");      
  }

  // view-source:http://oss.sheetjs.com/js-xlsx/ 
  function fixdata(data) {
    var o = "", l = 0, w = 10240;
    for(; l<data.byteLength/w; ++l) o+=String.fromCharCode.apply(null,new Uint8Array(data.slice(l*w,l*w+w)));
    o+=String.fromCharCode.apply(null, new Uint8Array(data.slice(l*w)));
    return o;
  }  
  // view-source:http://sheetjs.com/demos/writexlsx.html
  function s2ab(s) {
    if(typeof ArrayBuffer !== 'undefined') {
      var buf = new ArrayBuffer(s.length);
      var view = new Uint8Array(buf);
      for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    } else {
      var buf = new Array(s.length);
      for (var i=0; i!=s.length; ++i) buf[i] = s.charCodeAt(i) & 0xFF;
      return buf;
    }
  } 