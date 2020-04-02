var map = L.map( 'map', {
  center: [20.0, 5.0],
  minZoom: 2,
  zoom: 2
})

L.tileLayer( 'http://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
  attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>',
  subdomains: ['a', 'b', 'c']
}).addTo( map )

var myURL = jQuery( 'script[src$="demo.js"]' ).attr( 'src' ).replace( 'demo.js', '' )

var myIcon = L.icon({
  iconUrl: myURL + 'images/pin24.png',
  iconRetinaUrl: myURL + 'images/pin48.png',
  iconSize: [29, 24],
  iconAnchor: [9, 21],
  popupAnchor: [0, -14]
})

/* set up XMLHttpRequest */
var url = "http://localhost/openstreetmap/maps/Auto  par département dernier mars 2019.xlsx";
var oReq = new XMLHttpRequest();
oReq.open("GET", url, true);
oReq.responseType = "arraybuffer";

oReq.onload = function(e) {
  var arraybuffer = oReq.response;

  /* convert data to binary string */
  var data = new Uint8Array(arraybuffer);
  var arr = new Array();
  for (var i = 0; i != data.length; ++i) arr[i] = String.fromCharCode(data[i]);
  var bstr = arr.join("");

  /* Call XLSX */
  var workbook = XLSX.read(bstr, {
    type: "binary"
  });

  /* DO SOMETHING WITH workbook HERE */
  var first_sheet_name = workbook.SheetNames[0];
  /* Get worksheet */
  var worksheet = workbook.Sheets[first_sheet_name];
  var markers = XLSX.utils.sheet_to_json(worksheet, {
    raw: true
  });
  console.log(markers);

  for ( var i=0; i < markers.length; ++i )
  {
    L.marker( [markers[i]['Latitude'], markers[i]['Longitude']], {icon: myIcon} )
      .bindPopup(`
        <b>Département : </b> `+markers[i].Département+`<br/>
        <b>Code Département : </b> `+markers[i]["Code Département"]+`<br/>
        <b>Commune : </b> `+markers[i].Commune+`<br/>
        <b>Année : </b> `+markers[i].Année+`<br/>
        <b>Trimestre : </b> `+markers[i].Trimestre+`<br/>
        `)
      .addTo( map );
  }

}

oReq.send();