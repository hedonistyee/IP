# IP
<html>

<head>
	<meta http-equiv="content-type" content="text/html;charset=utf-8" />
	<script src="libraries/jquery.min.js"></script>
	<script src="libraries/papaparse.min.js"></script>
	<script src="libraries/xlsx.core.min.js"></script>
	<script src="libraries/bootstrap.min.js"></script>
	<link rel="stylesheet" href="libraries/bootstrap.min.css">
	<link rel="stylesheet" href="libraries/bootstrap-theme.min.css">
	<script src="libraries/bootstrap.file-input.js"></script>

	<script>
		function xlsxToJson(workbook, sheetName) {
			var csv = XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName]);
			var results = Papa.parse(csv, {
				delimiter: ",",
				header: true,
				skipEmptyLines: true,
				dynamicTyping: true
			});
			console.log(results);
			if (results.errors.length !== 0) {
				console.log(results.errors);
				if (results.errors[0].code != "UndetectableDelimiter") {
					alert("Something doesn't work with this file");
					throw new Error(results.error);
				}
			}
			return (results.data);
		}
	</script>

	<script>
		function doSomething(json) {
			$('#sheetJson').text(JSON.stringify(json));
			$('#containerSheetJson').show();
		}
	</script>

	<script>
		$(document).ready(function() {

			$('input[type=file]').bootstrapFileInput();

			var re = /(?:\.([^.]+))?$/; // to get file extension

			$("#xlsxfile").on("change", function() {
				var file = $("#xlsxfile")[0].files[0];
				var extension = re.exec(file.name)[1];
				if (extension != "csv") {
					var reader = new FileReader();
					reader.onload = function(e) {
						var workbook = null;
						try {
							workbook = XLSX.read(reader.result, {
								type: 'binary'
							});
						} catch (err) {
							alert("Something is wrong with this xlsx file");
							throw new Error(err);
						}
						// create the dropdown list
						var sheetNames = workbook.SheetNames;
						var size = 0;
						if (sheetNames.length < 3) {
							size = sheetNames.length + 1;
						} else {
							size = 4;
						}
						$('#selectSheet').empty();
						$('#selectSheet').append($("<option disabled selected>").attr('value', "").text("Select..."));
						$('#selectSheet').attr('size', size);
						$(sheetNames).each(function(index, item) {
							$('#selectSheet').append($("<option>").attr('value', index).text(item));
						});
						$('#containerSelectSheet').show();
						// submit action
						// $('#submitSheet').click(function() {
						// 	var json = xlsxToJson(workbook, $("#selectSheet option:selected").text());
						// 	// DO SOMETHING HERE
						// 	doSomething(json);
						// });
						// on select action
						$('#selectSheet').change(function() {
							var json = xlsxToJson(workbook, $("#selectSheet option:selected").text());
							// DO SOMETHING HERE
							doSomething(json);
						});
					};
					reader.onerror = function(err) {
						alert("I can't read this file!");
						throw new Error(err);
					};
					reader.readAsBinaryString(file);
				} else {
					$('#containerSelectSheet').hide();
					Papa.parse(file, {
						delimiter: ",",
						header: true,
						skipEmptyLines: true,
						dynamicTyping: true,
						complete: function(results) {
							var json = results.data;
							console.log(results);
							// DO SOMETHING HERE
							doSomething(json);
						}
					});
				}
			});
		});
	</script>

</head>

<body>

	<div class="container-fluid">
		<form id="xlsxform" method="post" enctype="multipart/form-data">
			<!-- <label for="xlsxfile" class="control-label">Select a XLSX file</label> -->
			<input id="xlsxfile" type="file" data-filename-placement="inside" class="btn-primary" title="xls(x) or csv" accept=".xlsx,.xls,.csv">
		</form>
	</div>

	<div class="container-fluid" id="containerSelectSheet" style="display:none">
		<div class="row">
			<div class="col-sm-3">
				<fieldset id="selectSheetFieldset">
					<label for="selectSheet" style="color:blue">Select a sheet</label>
					<select class="form-control" id="selectSheet" style="overflow-y:auto"></select>
				</fieldset>
			</div>
			<!-- <div class="col-sm-3">
				<button id="submitSheet" type="button" class="btn-primary">Submit</button>
			</div> -->
			<div class="col-sm-9"></div>
		</div>
	</div>

	<div class="container" id="containerSheetJson" style="display:none">
		<p><h3>Your sheet in JSON:</h3></p>
		<p id="sheetJson"></p>
	</div>
	<br>

</body>

</html>
