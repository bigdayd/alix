<?php

require __DIR__ . '/../vendor/autoload.php';

use Curl\Curl;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


$token = 'Test-4247c-3fc638-4b18-a01ed-1264T';
$login = '08022019@aliexpress-op.ru';
$key = $login . ':' . md5($login . gmdate('dmYH') . $token);


$mainColumns = [
    'product_id' => 'ID',
    'category_id' => 'Category ID',
    'subject' => 'Name',
];
$skuColumns = [
    'barcode' => 'Barcode',
    'sku_price' => 'Price',
    'sku_discount_price' => 'Discount Price',
    'currency_code' => 'Currency',
    'sku_stock' => 'Stock',
    'sku_code' => 'sku_code',
];
$propertiesColumns = [];


// get list
$curl = new Curl();
$curl->setHeader('Authorization', 'AccessToken '.$key);
$curl->post('https://alix.brand.company/api_top3/', array(
    'method' => 'aliexpress.solution.product.list.get',
    'aeop_a_e_product_list_query' => json_encode([
            'product_status_type'=> 'onSelling',
    ]),
));

if ($curl->error) {
    throw new Exception('Error: ' . $curl->errorCode . ': ' . $curl->errorMessage);
}

// get items
$products = [];
$result = json_decode($curl->response, true);
if ($result===false || !isset($result['aliexpress_solution_product_list_get_response']['result'])) {
    throw new Exception('Error: missing list response');
}
$list = $result['aliexpress_solution_product_list_get_response']['result']['aeop_a_e_product_display_d_t_o_list']['item_display_dto'];


foreach ($list as $item) {
    // get product
    $curl->post('https://alix.brand.company/api_top3/', array(
        'method' => 'aliexpress.solution.product.info.get',
        'product_id' => $item['product_id'],
    ));
    if ($curl->error) {
        throw new Exception('Error: ' . $curl->errorCode . ': ' . $curl->errorMessage);
    }

    $result = json_decode($curl->response, true);
    if ($result===false || !isset($result['aliexpress_solution_product_info_get_response']['result'])) {
        throw new Exception('Error: missing product response');
    }
    $data = $result['aliexpress_solution_product_info_get_response']['result'];
    $properties = $data['aeop_ae_product_propertys']['global_aeop_ae_product_property'];
    $skus = $data['aeop_ae_product_s_k_us']['global_aeop_ae_product_sku'];

    // main
    $row = [
        'product_id' => $data['product_id'],
        'category_id' => $data['category_id'],
        'subject' => $data['subject'],
        'sku' => [],
        'properties' => [],
    ];

    // sku
    foreach (array_keys($skuColumns) as $k) {
        $row['sku'][$k] = [];
    }
    foreach ($skus as $sku) {
        foreach (array_keys($skuColumns) as $k) {
            if (isset($sku[$k])) $row['sku'][$k][] = $sku[$k];
        }
    }
    foreach (array_keys($skuColumns) as $k) {
        $row['sku'][$k] = implode(';', $row['sku'][$k]);
    }

    // properties
    foreach ($properties as $property) {
        if (!isset($propertiesColumns[$property['attr_name_id']])) {
            $propertiesColumns[$property['attr_name_id']] = $property['attr_name'];
        }
        $row['properties'][$property['attr_name_id']] = isset($property['attr_value']) ? $property['attr_value'] : $property['attr_value_id'];
    }
    $products[] = $row;
}

unset($curl);

// prepare excel
$body = [array_merge(array_values($mainColumns), array_values($skuColumns), array_values($propertiesColumns))];
foreach ($products as $product) {
    $main = [];
    foreach (array_keys($mainColumns) as $k => $v) {
        $main[$k] = $product[$v];
    }
    $sku = [];
    foreach (array_keys($skuColumns) as $k => $v) {
        $sku[$k] = null;
        if (isset($product['sku'][$v])) {
            $sku[$k] = $product['sku'][$v];
        }
    }
    $properties = [];
    foreach (array_keys($propertiesColumns) as $k => $v) {
        $properties[$k] = null;
        if (isset($product['properties'][$v])) {
            $properties[$k] = $product['properties'][$v];
        }
    }
    $body[] = array_merge($main, $sku, $properties);
}

$spreadsheet = new Spreadsheet();
$spreadsheet->getActiveSheet()
    ->fromArray(
        $body,
        NULL,
        'A1'
    );

$writer = new Xlsx($spreadsheet);
$writer->save('products.xlsx');