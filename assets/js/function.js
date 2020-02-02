(function($){
	var lines = 0;
	var upload_status = "";
	var upload_msg = "";
	var fields = [];
	var fileName = '';


	function showBlockUI(msg){
		var message = 'Please Wait...';
		if (msg)
			message = msg;
	    $.blockUI({ 
	    	message: message,
	    	css: { 
	            border: 'none', 
	            padding: '15px', 
	            backgroundColor: '#000', 
	            '-webkit-border-radius': '10px', 
	            '-moz-border-radius': '10px', 
	            opacity: .5, 
	            color: '#fff' 
	    	} 
		}); 
	}
	function hideBlockUI(){
		$.unblockUI();
	}
	function uploadProduct(idx){
		showBlockUI('Uploading product ' + (idx-1));
		$.ajax({
			url: ecu_ajax_object.ajax_url,
			type: "POST",
			contentType: "application/x-www-form-urlencoded; charset=UTF-8",
			data: {
				action: "product_upload_action",
				fileName: fileName,
				lines: lines,
				idx: idx,
				fields: fields,
				nonce: ecu_ajax_object.nonce
			},
			success: function(result,status,xhr){
				// console.log(result);
				$("#excel-content").html(result);
				// alert(result);
				$("#btn-product-upload").attr("disabled","true");
				if(idx < lines){
					idx++;
					hideBlockUI();
					uploadProduct(idx)
				}else{
					hideBlockUI();
					$("#excel-content").html("<div>All products uploaded</div>");
				}
			},
			error: function(jqXHR, exception) {
				console.log("Failed");
				$("#excel-content").append("<div>" + (idx-1) + " Upload Failed</div>");
				hideBlockUI();
			}
		});		
	}
	$(document).ready(function () {
		// body...
		//uploaded excel file name
		
		//product upload button click
		
		$("#btn-product-upload").on("click", function(){
			console.log("upload started");
			
			//first row is header so omit it
			uploadProduct(2);
			// hideBlockUI();

		});
		//upload excel
		Dropzone.options.uploader = {
			maxFiles: 1,
			accept: function(file, done) {
				console.log("uploaded");
				done();

			},
			acceptedFiles: '.xlsx,.xls',
			init: function() {
				this.on("maxfilesexceeded", function(file){
					this.hiddenFileInput.removeAttribute('multiple');
					this.removeAllFiles(); 
					this.addFile(file); 
				});
				this.on("error", function(e){
					var error = e.previewElement.getElementsByClassName("dz-error-message")[0].innerText;
					Swal.fire({
					  type: 'error',
					  title: 'Uploading failed...',
					  text: error,
					});
					this.removeAllFiles();
					//disable upload button
					$("#btn-product-upload").attr("disabled","true");
				});
                this.on("addedfile", function(e) {
                    var removeBtn = Dropzone.createElement("<a href='javascript:;' class='btn red btn-sm btn-block'>Remove</a>");
                    var target = this;
                    removeBtn.addEventListener("click", function(n) {
                        n.preventDefault();
                        n.stopPropagation();
                        target.removeFile(e);
                        console.log("deleting...");
                        //disable upload button
                        showBlockUI();
                        $("#btn-product-upload").attr("disabled","true");
    					$.ajax({
							url: ecu_ajax_object.ajax_url,
							type: "POST",
							contentType: "application/x-www-form-urlencoded; charset=UTF-8",
							data: {
								action: "deleteaction",
								fileName: fileName,
								nonce: ecu_ajax_object.nonce
							},
							success: function(data){
								console.log(data);

								$("#excel-content").empty();
								hideBlockUI();
							},
							error: function(jqXHR, exception) {
								console.log("Failed");
								$("#excel-content").append("<div>Upload Failed</div>");
								hideBlockUI();
							}
						});
                    });
                    e.previewElement.appendChild(removeBtn);
                });
                this.on("success", function(data){
					// Swal.fire(
					//   'Uploaded successfully',
					//   '',
					//   'success'
					// )
					//enable upload button
					
					fileName = data.upload.filename;
					$.ajax({
						url: ecu_ajax_object.ajax_url,
						type: "POST",
						contentType: "application/x-www-form-urlencoded; charset=UTF-8",
						data: {
							action: "get_excel_info",
							fileName: fileName,
							nonce: ecu_ajax_object.nonce
						},
						success: function(data){
							var info = JSON.parse(data);
							lines = info.count;
							upload_msg = info.message;
							upload_status = info.status;
							fields = info.fields;
							console.log(info);
							hideBlockUI();
							$("#btn-product-upload").removeAttr("disabled");
						},
						error: function(jqXHR, exception) {
							console.log("Failed");
							$("#excel-content").append("<div>Upload Failed</div>");
							hideBlockUI();
						}
					});

                });
			}
		};
	})
})(jQuery);