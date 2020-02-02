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
	if( !isset($allDataInSheet) )
		wp_die("Dataset is empty!");
	
	$post_title = sanitize_text_field($allDataInSheet[$i][$fieldIDs['post_title']]);

	if(post_exists($post_title)===0){
		wp_die("Post does not exist!");
	}
	else{
		$id = post_exists($post_title);
		ob_start();
		echo "<div>Update " . ($i-1) ."th product</div>";
		if(isset($allDataInSheet[$i][$fieldIDs["_sku"]])){
			$cellValue = $allDataInSheet[$i][$fieldIDs["_sku"]];					
			if ( !$cellValue && !empty($fieldIDs["_sku"]) ) {
				$cellValue = '';
			}
			else 
				update_post_meta( $id, "_sku", $cellValue );						
		}
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