<?php
/*
 * Plugin Name: ECU Product Importer for WooCommerce 
 * Description: Import ECU proudcts with Excel
 * Version: 1.0
 * Author: Shao
 *
 * WC requires at least: 2.2
 * WC tested up to: 3.6.5
 *  
 * License: GPL2
 * Created On: 12-09-2019
 * Updated On: 18-09-2019
 */

if ( ! defined( 'ABSPATH' ) ) exit; // Exit if accessed directly
global $allDataInSheet;
global $fieldIDs;
require_once( ABSPATH . 'wp-admin/includes/image.php' );
//ADD MENU LINK AND PAGE FOR WOOCOMMERCE IMPORTER
add_action('admin_menu', 'ecu_product_importer_menu');

function ecu_product_importer_menu() {
	// add_submenu_page( 'edit.php?post_type=product', 'Product Excel Importer & Bulk Editing', 'Product Excel Importer & Bulk Editing', 'manage_options', 'webd-woocommerce-product-excel-importer-bulk-edit', 'webd_woocommerce_product_excel_importer_bulk_edit_init' );	
	// add_submenu_page( 'woocommerce', 'Product Excel Importer & Bulk Editing', 'Product Excel Importer & Bulk Editing', 'manage_options', 'webd-woocommerce-product-excel-importer-bulk-edit', 'webd_woocommerce_product_excel_importer_bulk_edit_init' );
	add_menu_page('ECU Product Excel Importer', 'ECU Product Importer', 'administrator', 'ecu-product-importer', 'ecu_product_importer_init', 'dashicons-products','50');
}
//load css and js
function ecu_product_importer_enqueue_scripts(){
	//excel upload css
	wp_register_style( 'ecu_product_importer_enqueue_dropzone_css', plugins_url( "/assets/css/dropzone.css", __FILE__ ) );
	wp_enqueue_style( 'ecu_product_importer_enqueue_dropzone_css');
	//alert css
	wp_register_style( 'ecu_product_importer_enqueue_sweetalert_css', plugins_url( "/assets/css/sweetalert.min.css", __FILE__ ) );
	wp_enqueue_style( 'ecu_product_importer_enqueue_sweetalert_css');
	//plugin css
	wp_register_style( 'ecu_product_importer_enqueue_css', plugins_url( "/assets/css/style.css", __FILE__ ) );
	wp_enqueue_style( 'ecu_product_importer_enqueue_css');
	//excel upload js
    wp_register_script( 'ecu_product_importer_enqueue_dropzone_js', 
    	plugins_url( '/assets/js/dropzone.js', __FILE__ ), 
    	array('jquery') , 
    	null, 
    	true
    );
    wp_enqueue_script( 'ecu_product_importer_enqueue_dropzone_js');
    //alert js
    wp_register_script( 'ecu_product_importer_enqueue_sweetalert_js', 
    	plugins_url( '/assets/js/sweetalert.min.js', __FILE__ ), 
    	array('jquery') , 
    	null, 
    	true
    );
    wp_enqueue_script( 'ecu_product_importer_enqueue_sweetalert_js');
    // load jquery block UI
    wp_register_script('ecu_product_importer_enqueue_block_ui_js',
		plugins_url( '/assets/js/jquery.blockUI.min.js', __FILE__ ),
		array('jquery') , 
		null,
		true
	);
	wp_enqueue_script('ecu_product_importer_enqueue_block_ui_js');
    //plugin js
    wp_register_script( 'ecu_product_importer_enqueue_js', 
    	plugins_url( '/assets/js/function.js', __FILE__ ), 
    	array('jquery', 'ecu_product_importer_enqueue_dropzone_js','ecu_product_importer_enqueue_sweetalert_js') , 
    	null, 
    	true
    );
    wp_enqueue_script( 'ecu_product_importer_enqueue_js');

    wp_localize_script('ecu_product_importer_enqueue_js', 
    	'ecu_ajax_object', 
    	array(
    		'plugin_url' => plugins_url( '', __FILE__ ),
    		'site_url' => site_url(),
    		'upload_url' => plugins_url( '/upload', __FILE__ ),
    		'nonce' => wp_create_nonce( 'ajax-nonce' ),
    		'ajax_url' => admin_url( 'admin-ajax.php' )
    	)
    );
}

//set global variables
function set_global_variable(){
	global $allDataInSheet;
	global $fieldIDs;
}
add_action( 'init', 'set_global_variable' );

add_action('admin_enqueue_scripts', 'ecu_product_importer_enqueue_scripts');
//get excel file information
add_action("wp_ajax_get_excel_info", "ecu_get_excel_file_info");
function ecu_get_excel_file_info(){
    $nonce = $_POST['nonce'];
    if ( ! wp_verify_nonce( $nonce, 'ajax-nonce' ) )
        wp_die ( 'Busted!');
    $excel_file = dirname(__FILE__) . "/upload/" . $_POST["fileName"];
    require_once( plugin_dir_path( __FILE__ ) .'/Classes/PHPExcel/IOFactory.php');

	try {
		$objPHPExcel = PHPExcel_IOFactory::load($excel_file);
	} catch(Exception $e) {
		die('Error loading file "'.pathinfo($excel_file,PATHINFO_BASENAME).'": '.$e->getMessage());
	}
	global $allDataInSheet;
	$allDataInSheet = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
	$data = count($allDataInSheet);  // Here get total count of row in that Excel sheet
		
	$rownumber=1;
	$row = $objPHPExcel->getActiveSheet()->getRowIterator($rownumber)->current();
	$cellIterator = $row->getCellIterator();
	$cellIterator->setIterateOnlyExistingCells(false);

	foreach ($cellIterator as $cell) {
		//getValue
		global $fieldIDs;
		$fieldIDs[sanitize_text_field($cell->getValue())] = sanitize_text_field($cell->getColumn());
	}
	if(!$fieldIDs["product_image1_url"]){
		$result["status"] = "fail";
		$result["message"] = "Excel format is wrong!";
		$result["fields"] = $fieldIDs["product_image1_url"];
		$result["count"] = 0;
		echo json_encode($result);
		wp_die();
	}

	// echo $data;
	$line_count = $data;
	$result["status"] = "success";
	$result["message"] = "Success";
	$result["fields"] = $fieldIDs;
	$result["count"] = $line_count;
	echo json_encode($result);
	wp_die();
}
//delete file - ajax reqeust
add_action( "wp_ajax_deleteaction", "ecu_delete_excel_file" );
function ecu_delete_excel_file(){
    // Check for nonce security
    $nonce = $_POST['nonce'];
    if ( ! wp_verify_nonce( $nonce, 'ajax-nonce' ) )
        wp_die ( 'Busted!');
	$excel_file = dirname(__FILE__) . "/upload/" . $_POST["fileName"];
	if( file_exists($excel_file)){
		unlink($excel_file);
		echo "File deleted";
		wp_die();
	}
	echo $excel_file;
	wp_die();
}
//product upload logic
add_action( "wp_ajax_product_upload_action", "ecu_product_upload" );
function ecu_product_upload(){
	if($_SERVER['REQUEST_METHOD'] != 'POST' || !current_user_can('administrator') )
		wp_die("You don't have enough permission!");

	$excel_file = dirname(__FILE__) . "/upload/" . $_POST["fileName"];
	require_once( plugin_dir_path( __FILE__ ) .'/Classes/PHPExcel/IOFactory.php');

	try {
		$objPHPExcel = PHPExcel_IOFactory::load($excel_file);
	} catch(Exception $e) {
		die('Error loading file "'.pathinfo($excel_file,PATHINFO_BASENAME).'": '.$e->getMessage());
	}
	$allDataInSheet = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
	$data = count($allDataInSheet);  // Here get total count of row in that Excel sheet
		
	$rownumber=1;
	$row = $objPHPExcel->getActiveSheet()->getRowIterator($rownumber)->current();
	$cellIterator = $row->getCellIterator();
	$cellIterator->setIterateOnlyExistingCells(false);

	$i = $_POST["idx"];
	$fieldIDs = $_POST["fields"];

	$sku = $allDataInSheet[$i][$fieldIDs['_sku']];
	if($allDataInSheet[$i][$fieldIDs['skip']] == "1"){
		echo "<div>{$sku} is skipped!</div>";
		wp_die();
	}

	if(!$fieldIDs["product_image1_url"])
		wp_die("Excel is not matched");

	if( !isset($allDataInSheet) )
		wp_die("Dataset is empty!");
	
	$post_title = sanitize_text_field($allDataInSheet[$i][$fieldIDs['post_title']]);
	$post_content = ($allDataInSheet[$i][$fieldIDs['post_content']]);
	$post_excerpt = ($allDataInSheet[$i][$fieldIDs['post_excerpt']]);
	$post_name = sanitize_title_with_dashes($allDataInSheet[$i][$fieldIDs['post_title']]);
	$post_type = 'product';


	if(post_exists($post_title)===0){
		$post = array(
			'post_title'   => $post_title,
			'post_content' => $post_content,
			'post_status'  => 'publish',
			'post_excerpt' => $post_excerpt,
			'post_name'    => $post_name,
			'post_type'    => $post_type
		);
		ob_start();
		echo "<div>Insert " . ($i-1) ."th product</div>";
		// wp_die();
		$id = wp_insert_post( $post);
	}
	else{

		$id = post_exists($post_title);
		// delete existing product images
		// $args = array( 
		//     'post_parent' => $id,
		//     'post_type' => 'attachment'
		// );
		// $attachments = get_posts( $args );
		// if (is_array($attachments) && count($attachments) > 0) {

		// 	// Delete all the Children of the Parent Page
		// 	foreach($attachments as $attachment){

		// 	    wp_delete_post($attachment->ID, true);

		// 	}
		// }
		$post = array(
			'ID' 		   => $id,
			'post_title'   => $post_title,
			'post_content' => $post_content,
			'post_status'  => 'publish',
			'post_excerpt' => $post_excerpt,
			'post_name'    => $post_name,
			'post_type'    => $post_type
		);
		ob_start();
		echo "<div>Update " . ($i-1) ."th product</div>";
		// wp_die();
		wp_update_post($post);
		// print "<p><a href='".esc_url( get_permalink($id))."' target='_blank'>".$title."</a> already exists. Updated.</p>";
	}
	// update product info
	if(isset($allDataInSheet[$i][$fieldIDs['_sale_price']])){
		$sale_price = sanitize_text_field($allDataInSheet[$i][$fieldIDs['_sale_price']]);					

		if ( strlen(trim($sale_price)) >=1 ) {
			update_post_meta( $id, '_sale_price', $sale_price );
		}
		else
			delete_post_meta($id,'_sale_price');
		
	}
	if(isset($allDataInSheet[$i][$fieldIDs['_regular_price']])){
		$regular_price = sanitize_text_field($allDataInSheet[$i][$fieldIDs['_regular_price']]);
		if ( !$regular_price  && !empty($allDataInSheet[$i][$fieldIDs['_regular_price']])) {
		  $regular_price = '';
		  // print "For regular price of {$post_title} you need numbers entered.<br/>";
		}else update_post_meta( $id, '_regular_price', $regular_price );						
	}
	//ADDITION : IF SALE PRICE IS EMPTY PRICE WILL BE EQUAL TO REGULAR PRICE
	if(strlen(trim($sale_price)) >=1){
		//echo ("sale price is not empty");
		update_post_meta( $id, '_price', $sale_price );			
	}elseif(isset($allDataInSheet[$i][$fieldIDs['_regular_price']])){
		// delete_post_meta($id,'_sale_price');
		//echo ("regular price is");
		update_post_meta( $id, '_price', $regular_price );
	}

	foreach($fieldIDs as $key => $value){
		switch ($key) {
			case "post_title":
			case "post_content":
			case "post_excerpt":
			case "post_title":
			case "post_type":
			case "product_image1_url":
			case "product_image1_alt_text":
			case "product_image1_title":
			case "product_image1_caption":
			case "product_image1_description":
			case "product_image2_url":
			case "product_image2_alt_text":
			case "product_image2_title":
			case "product_image2_caption":
			case "product_image2_description":
			case "product_image3_url":
			case "product_image3_alt_text":
			case "product_image3_title":
			case "product_image3_caption":
			case "product_image3_description":
			case "product_image4_url":
			case "product_image4_alt_text":
			case "product_image4_title":
			case "product_image4_caption":
			case "product_image4_description":
			case "product_image5_url":
			case "product_image5_alt_text":
			case "product_image5_title":
			case "product_image5_caption":
			case "product_image5_description":
			case "category":
			case "_regular_price":
			case "_sale_price":
			case "share":
			case "product_tags":
			case "skip":
				// continue;
				break;
			default:
				# code...
				if(isset($allDataInSheet[$i][$fieldIDs[$key]])){
					$cellValue = $allDataInSheet[$i][$fieldIDs[$key]];					
					if ( !$cellValue && !empty($fieldIDs[$key]) ) {
						$cellValue = '';
						// print "For {$key} of {$post_title} you need numbers entered.<br/>";
					}
					else 
						update_post_meta( $id, $key, $cellValue );						
				}
				break;
		}
	}
	//path to images
	$upload_dir = wp_upload_dir();
	$upload_dir = $upload_dir['baseurl'];
	$upload_dir = $upload_dir . '/products';

	$product_gallery = array();
	//set pictures
	for($idx=1;$idx<=5;$idx++){
		// print "product_image{$idx}_url";
		if( isset($allDataInSheet[$i][$fieldIDs["product_image{$idx}_url"]]) && strlen($allDataInSheet[$i][$fieldIDs["product_image{$idx}_url"]]) > 0 ){
			$product_image1_url = ($allDataInSheet[$i][$fieldIDs["product_image{$idx}_url"]]);
			$ext = pathinfo($product_image1_url, PATHINFO_EXTENSION);
			$product_image_title = ($allDataInSheet[$i][$fieldIDs["product_image{$idx}_title"]]);
			// if (strpos($product_image1_url, " ") !== false) {
		    $product_image1_url_nospace = str_replace(" ", "-", $product_image_title) . '.' . $ext;
		    copy(wp_upload_dir()["basedir"] . "/source/" . $product_image1_url, wp_upload_dir()["basedir"] . "/products/" . $product_image1_url_nospace);
			    // print("renamed!!!!");
			// }
			// else
				// $product_image1_url_nospace = $product_image1_url;
			$product_image1_alt_text = sanitize_text_field($allDataInSheet[$i][$fieldIDs["product_image{$idx}_alt_text"]]);
			$product_image1_title = sanitize_text_field($allDataInSheet[$i][$fieldIDs["product_image{$idx}_title"]]);
			$product_image1_caption = sanitize_text_field($allDataInSheet[$i][$fieldIDs["product_image{$idx}_caption"]]);
			$product_image1_description = sanitize_text_field($allDataInSheet[$i][$fieldIDs["product_image{$idx}_description"]]);
			$file_path = $upload_dir . "/" . $product_image1_url_nospace;
			$file_type = wp_check_filetype( basename( $file_path ), null );

			if(post_exists($product_title)===0){
				$attachment = array(
					"guid"			 => $file_path,
				    "post_mime_type" => $file_type["type"],
				    "post_title"     => $product_image_title,
				    "post_content"   => $product_image1_description,
				    "post_excerpt"   => $product_image1_caption,
				    "post_type"		 => "attachment",
				    "post_status"    => "inherit"
				);
				$attach_id = wp_insert_attachment( $attachment, wp_upload_dir()["basedir"] . "/products/" . $product_image1_url_nospace, $id );
								// Generate the metadata for the attachment, and update the database record.
				$attach_data = wp_generate_attachment_metadata( $attach_id, wp_upload_dir()["basedir"] . "/products/" . $product_image1_url_nospace );
			}
			else{
				$attach_id =  post_exists($product_title);
				$attachment = array(
					"ID"			 => $attach_id,
					"guid"			 => $file_path,
				    "post_mime_type" => $file_type["type"],
				    "post_title"     => $product_image_title,
				    "post_content"   => $product_image1_description,
				    "post_excerpt"   => $product_image1_caption,
				    "post_type"		 => "attachment",
				    "post_status"    => "inherit"
				);
				wp_update_post($attachment);
			}
			// Insert the attachment.
			wp_update_attachment_metadata( $attach_id,  $attach_data );
			update_post_meta( $attach_id, "_wp_attachment_image_alt", $product_image1_alt_text );
			if($idx == 1)
				set_post_thumbnail( $id, $attach_id );
			else
				array_push($product_gallery, $attach_id);

			//remove product cache data
			// wc_delete_product_transients( $id );	
		}
	}
	//product image gallery
	update_post_meta( $id, "_product_image_gallery", implode(",",$product_gallery));

	//set product tags
	$product_tags = explode(',',sanitize_text_field($allDataInSheet[$i][$fieldIDs['product_tags']]));
	foreach($product_tags as $tag){
		if(term_exists($tag, "product_tag") === null){
			wp_insert_term($tag,"product_tag");
		}
		// else{
		// 	$term_id = term_exists($tag);
		// 	wp_update_term($tag,"product_tag");
		// }
	}
	wp_set_object_terms($id, $product_tags, "product_tag"); 

	//set product categories
	$product_cats = explode(',',sanitize_text_field($allDataInSheet[$i][$fieldIDs['category']]));
	
	foreach ($product_cats as $cats) {
		$parent_cat_id = 0;
		$cat_array = explode('/', $cats);
		$cat_ids = array();
		foreach( $cat_array as $cat){
			if(term_exists($cat, "product_cat", $parent_cat_id) === null){
				$tmp = wp_insert_term($cat,"product_cat", array("parent" => $parent_cat_id));
				$parent_cat_id = $tmp["term_id"];
				array_push($cat_ids, $parent_cat_id);
			}
			else{
				$term_id = term_exists($cat, "product_cat", $parent_cat_id);
				$tmp = wp_update_term($term_id["term_id"],"product_cat", array("parent" => $parent_cat_id));
				$parent_cat_id = $tmp["term_id"];
				array_push($cat_ids, $parent_cat_id);
			}
		}
		wp_set_object_terms($id, $cat_ids, "product_cat"); 
		# code...
	}
	wp_die(" Success!");
}
//init admin page
function ecu_product_importer_init(){
	?>
		<div class="container">
		    <div class="row h-center">
			    <div class="col-md-12">
			        <form method="post" 
			        	enctype="multipart/form-data" 
			        	action="<?php echo plugins_url('/upload-excel.php', __FILE__) ?>" 
			        	class="dropzone dropzone-file-area" 
			        	id="uploader"
			        >
		            <div class="dz-message" data-dz-message><h3><span>Drop a file here or click to upload</span></h3></div>
			        </form>

			    </div>
			</div>
			<div class="row h-center">
			    <div class="col-md-12">
			    	<button id="btn-product-upload" type="button" class="btn btn-lg btn-primary" disabled>Upload Products</button>
			    </div>
			</div>
			<div class="row h-center">
				<div id="excel-content" class="col-md-12">
				</div>
			</div>
		</div>;
	<?php
}
?>