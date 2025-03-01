<?php
/*
Plugin Name: Export Customer Phone Numbers to Excel
Description: این افزونه شماره تماس تمام مشتریانی که سفارش داده‌اند را به صورت فایل اکسل خروجی می‌گیرد.
Version: 1.0
Author: SajjadAtaei
*/

if (!defined('ABSPATH')) {
    exit;
}

add_action('admin_menu', 'ecpn_add_admin_menu');

function ecpn_add_admin_menu() {
    add_menu_page(
        'خروجی اکسل مشتریان', 
        'خروجی اکسل مشتریان', 
        'manage_options', 
        'export-customer-phone-numbers', 
        'ecpn_display_export_page', 
        'dashicons-phone', 
        6 //
    );
}


function ecpn_display_export_page() {
    if (!current_user_can('manage_options')) {
        return;
    }

    echo '<div class="wrap">';
    echo '<h1>خروجی اکسل مشتریان</h1>';
    echo '<p>برای دریافت فایل اکسل شماره تماس مشتریان، روی دکمه زیر کلیک کنید.</p>';
    echo '<a href="' . admin_url('admin.php?page=export-customer-phone-numbers&ecpn_export_excel=1') . '" class="button button-primary">دریافت فایل اکسل</a>';
    echo '</div>';
}


add_action('admin_init', 'ecpn_export_excel');

function ecpn_export_excel() {
    if (isset($_GET['ecpn_export_excel']) && $_GET['ecpn_export_excel'] == '1') {
        if (!current_user_can('manage_options')) {
            return;
        }

        
        $orders = wc_get_orders(array(
            'limit' => -1, 
            'status' => 'completed', 
        ));

        
        require_once plugin_dir_path(__FILE__) . 'vendor/autoload.php';

         
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        
        $sheet->setCellValue('A1', 'نام مشتری');
        $sheet->setCellValue('B1', 'شماره تماس');
        $sheet->setCellValue('C1', 'ایمیل');

        
        $row = 2;
        foreach ($orders as $order) {
            $customer_name = $order->get_billing_first_name() . ' ' . $order->get_billing_last_name();
            $customer_phone = $order->get_billing_phone();
            $customer_email = $order->get_billing_email();

            $sheet->setCellValue('A' . $row, $customer_name);
            $sheet->setCellValue('B' . $row, $customer_phone);
            $sheet->setCellValue('C' . $row, $customer_email);
            $row++;
        }

        
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="customer_phones.xlsx"');
        $writer->save('php://output');
        exit;
    }
}