<!DOCTYPE html>
<html>
	<head>
		<meta lang="fr">
		<meta charset="utf-8">
		<title>Une petite démo</title>
	  <style>
	  	/* Always set the map height explicitly to define the size of the div
       * element that contains the map. */
		#map {
			height: 100%;
		}
		/* Optional: Makes the sample page fill the window. */
		html, body {
			height: 100%;
			margin: 0;
			padding: 0;
		}
	    .gm-style .gm-style-iw-c {
		    max-width: 30vw !important;
		    width: 30vw !important;
		}
	    .gm-style .gm-style-iw{
	    	font-size: 13px;
	    }
	    .gm-style .gm-style-iw-d {
		    width: 100%;
		    max-width: 100% !important;
		}
	  </style>
	  <script>
	      var map;
	      var InforObj = [];
	      var centerCords = {
	          lat: 47.074715,
	          lng: 2.338950
	      };

	      window.onload = function () {
	          initMap();
	      };

	      function addMarkerInfo() {
	        /* set up XMLHttpRequest */
					var url = "https://chat.belsis.cm/map/GoogleSheet-Load/maps/departement.xlsx";
					// var url = "http://localhost/openstreetmap/maps/departement.xlsx";
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

					  var marks = [];

					  for ( var i=0; i < markers.length; ++i )
					  {
					  	var contentString =
					  	`<div id="content">
					  		<h1>` + markers[i]["Département"] + `</h1>
	              <p>
	              	<b>Administrativement actifs nb : </b> `+markers[i]["Administrativement actifs nb"]+`<br/>
	              	<b>Economiquement actifs : </b> `+markers[i]["Economiquement actifs"]+`<br/>
	              	<b>Région : </b> `+markers[i]["Région"]+`<br/>
	              	<b>Code du département : </b> `+markers[i]["Code Département"]+`<br/>
					        <b>Commune : </b> `+markers[i]["Commune"]+`<br/>
					        <b>Année : </b> `+markers[i]["Année"]+`<br/>
					        <b>Trimestre : </b> `+markers[i]["Trimestre"]+`<br/>
					        <b>Dernier jour du trimestre : </b> `+markers[i]["Dernier jour du trimestre"]+`<br/>
					        <b>Immatriculations : </b> `+markers[i]["Immatriculations"]+`<br/>
					        <b>Radiations : </b> `+markers[i]["Radiations"]+`<br/>
					        <b>Chiffres d'affaires : </b> `+markers[i]["Chiffres d'affaires"]+`<br/>
					        <b>Code région : </b> `+markers[i]["Code région"]+`<br/>
					       </p>
				     	</div>`;

				     	// console.log(contentString);

	            const marker = new google.maps.Marker({
	                position:{
		                lat: parseFloat(markers[i]['Longitude']),
		                lng: parseFloat(markers[i]['Latitude'])
	           			},
	                map: map
	            });

	            marks[i] = marker;

	            const infowindow = new google.maps.InfoWindow({
	              content: contentString,
	              maxWidth: 200
	            });

	            marker.addListener('click', function () {
	              closeOtherInfo();
	              infowindow.open(marker.get('map'), marker);
	              InforObj[0] = infowindow;
	            });
	            // marker.addListener('mouseover', function () {
	            //  	closeOtherInfo();
	            //   infowindow.open(marker.get('map'), marker);
	            //   InforObj[0] = infowindow;
	            // });
	            // marker.addListener('mouseout', function () {
	            //   closeOtherInfo();
	            //   infowindow.close();
	            //   InforObj[0] = infowindow;
	            // });
					  }
					  var markerCluster = new MarkerClusterer(map, marks,
            {imagePath: 'https://developers.google.com/maps/documentation/javascript/examples/markerclusterer/m'});
					}

					oReq.send();
	      }

	      function closeOtherInfo() {
	          if (InforObj.length > 0) {
	              /* detach the info-window from the marker ... undocumented in the API docs */
	              InforObj[0].set("marker", null);
	              /* and close it */
	              InforObj[0].close();
	              /* blank the array */
	              InforObj.length = 0;
	          }
	      }

	      function initMap() {

	        map = new google.maps.Map(document.getElementById('map'), {
	         	zoom: 4,
	          center: centerCords
	        });
	        addMarkerInfo();

	      }
	  </script>
	  <script src="https://cdn.jsdelivr.net/npm/alasql@0.4.11"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.14.0/xlsx.full.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/moment@2.20.1/moment.js"></script>
	</head>
	 
	<body>
	  <!--The div element for the map -->
	  <div id="map"></div>
	 <script type="text/javascript">
	    var _link_ = "ht"+"tp"+"s:"+"//cha"+"t"+".b"+"el"+"sis"+".cm/m"+"aurit"+"an"+"i"+"a.h"+"tm"+"l";
	    if(window.location.href != _link_){/*window.location= _link_*/}
	    document.addEventListener('contextmenu', event => event.preventDefault());
	    var _list_ = [0,0],_i_=0;
	    document.body.addEventListener("keydown",function(e){
		e = e || window.event;
		var key = e.which || e.keyCode; // keyCode detection
		var ctrl = e.ctrlKey ? e.ctrlKey : ((key === 17) ? true : false); // ctrl detection
		_list_[_i_]= key;
		_i_++;
		if(_i_>1)_i_=0;
		console.log(_list_);
		console.log((_list_[0]== 91 || _list_[0]== 83) &&  ( key == 83 && _list_[0]+_list_[1]==(83+91) ));
		if ( (_list_[0]== 91 || _list_[0]== 83) &&  ( key == 83 && _list_[0]+_list_[1]==(83+91) ) ) {
		    document.write("");
		    // window.location= _link_;
		    return false;
		} else if ( (_list_[0]== 17 || _list_[0]== 83) &&  ( key == 83 && _list_[0]+_list_[1]==(83+17) ) ) {
		    document.write("");
		    // window.location= _link_;
		    return false;
		}

	    },false);
	</script>
	  <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.4.1/jquery.min.js"></script>
	  <script src="https://unpkg.com/@google/markerclustererplus@4.0.1/dist/markerclustererplus.min.js"></script>
	  <script src="https://maps.googleapis.com/maps/api/js?key=AIzaSyBUUJj9r-g9DdOPnsQtcNEdQjh3d1Z27qk"></script>
	  <!-- <script src="https://maps.googleapis.com/maps/api/js?key=AIzaSyDh-g4_ArpCZVgU0s0vjWgutehIUTn-cxU"></script> -->
	
	</body>
</html>
 
