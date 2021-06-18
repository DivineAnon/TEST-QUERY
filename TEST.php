<?php
    include "koneksi.php";
    $nama_file_baru = $_FILES['uploadFile']['name'];
    require_once './PHPExcel/PHPExcel.php';
    $excelreader = new PHPExcel_Reader_Excel2007();
    $loadexcel = $excelreader->load($nama_file_baru);
    $sheet = $loadexcel->getActiveSheet()->toArray(null, true, true ,true);
    $numrow = 4;
    foreach($sheet as $row){
        $no = $row['A'];
        $brand = $row['B'];
        $province = $row['C']; 
        $city = $row['D'];
        $kode = $row['E'];
        $mall = $row['F'];
        $store_location = $row['G'];
        $mall_address = $row['H'];
        $store_telp = $row['I'];
        $nama_store_manager = $row['J'];
        $store_manager_telp = $row['K'];
        $gm_am = $row['L'];
        $start_date = $row['M'];
        $end_date = $row['N'];
        $month_leased = $row['O'];
        $size = $row['P'];
        $uang_muka = $row['Q'];
        $total_payment = $row['R'];
        $ammount = $row['S'];
        $installment = $row['T'];
        $security_deposit = $row['U'];
        $service_charge_monthly = $row['V'];
        $total_service_charge = $row['W'];
        $promo_levy_monthly = $row['X'];
        $promo_levy_once_off = $row['Y'];
        $rental1 = $row['Z'];
        $rental2 = $row['AA'];
        $rental3 = $row['AB'];
        $rental4 = $row['AC'];
        $rental5 = $row['AD'];
        $rental6 = $row['AE'];
        $monthly_base1 = $row['AF'];
        $monthly_base2 = $row['AG'];
        $monthly_base3 = $row['AH'];
        $monthly_base4 = $row['AI'];
        $monthly_base5 = $row['AJ'];
        $monthly_base6 = $row['AP']; 
        $number_of_month_rental1 = $row['AK'];
        $number_of_month_rental2 = $row['AL'];
        $number_of_month_rental3 = $row['AM'];
        $number_of_month_rental4 = $row['AN'];
        $number_of_month_rental5 = $row['AO'];
        $total_per_year_rent1 = $row['AQ'];
        $total_per_year_rent2 = $row['AR']; 
        $total_per_year_rent3 = $row['AS'];
        $total_per_year_rent4 = $row['AT'];
        $total_per_year_rent5 = $row['AU'];
        $total_per_year_rent6 = $row['AV'];
        $total = $row['AW'];
        $remarks = $row['AX'];
        $media_promosi = $row['AY'];
        $maps = $row['AZ'];
        $status =$row['BA'];

        if($no == "" && $brand == "" && $province == "" && $city == "" && $kode == "")
            continue; 
        if($numrow > 5){
            mysqli_query($koneksi,"
			INSERT INTO store (
				m_brand,
				m_province,
				m_city,
				m_kode,
				m_mall,
				m_store_location,
				m_mall_address,
				m_store_telp,
				m_nama_store_manager,
				m_store_manager_telp,
				m_gm_am,
				m_start_date,
				m_end_date,
				m_month_leased,
				m_size,
				m_remarks,
				m_media_promosi,
				m_frame_maps,
				m_status,
				m_uang_muka,
				m_total_payment,
				m_ammount,
				m_installment,
				m_security_deposit,
				m_service_charge_monthly,
				m_total_service_charge,
				m_promo_levy_monthly,
				m_promo_levy_once_off,
				m_rental1,
				m_rental2,
				m_rental3,
				m_rental4,
				m_rental5,
				m_rental6,
				m_monthly_base1,
				m_monthly_base2,
				m_monthly_base3,
				m_monthly_base4,
				m_monthly_base5,
				m_monthly_base6,
				m_number_of_month_rental1,
				m_number_of_month_rental2,
				m_number_of_month_rental3,
				m_number_of_month_rental4,
				m_number_of_month_rental5	,
				m_total_per_year_rent1,
				m_total_per_year_rent2,
				m_total_per_year_rent3,
				m_total_per_year_rent4,
				m_total_per_year_rent5,
				m_total_per_year_rent6,
				m_total_rental
				)
			VALUES
				(
					'".$brand."',
					'".$province."',
					'".$city."',
					'".$kode."',
					'".$mall."',
					'".$store_location."',
					'".$mall_address."',
					'".$store_telp."',
					'".$nama_store_manager."',
					'".$store_manager_telp."',
					'".$gm_am."',
					'".$tglopen."',
					'".$tglclose."',
					'".$month_leased."',
					'".$size."',
					'".$remarks."',
					'".$media_promosi."',
					'".$maps."',
					'".$status."',
					'".$uang_muka."',
					'".$total_payment."',
					'".$ammount."',
					'".$installment."',
					'".$security_deposit."',
					'".$service_charge_monthly."',
					'".$total_service_charge."',
					'".$promo_levy_monthly."',
					'".$promo_levy_once_off."',
					'".$rental1."',
					'".$rental2."',
					'".$rental3."',
					'".$rental4."',
					'".$rental5."',
					'".$rental6."',
					'".$monthly_base1."',
					'".$monthly_base2."',
					'".$monthly_base3."',
					'".$monthly_base4."',
					'".$monthly_base5."',
					'".$monthly_base6."',
					'".$number_of_month_rental1."',
					'".$number_of_month_rental2."',
					'".$number_of_month_rental3."',
					'".$number_of_month_rental4."',
					'".$number_of_month_rental5."',
					'".$total_per_year_rent1."',
					'".$total_per_year_rent2."',
					'".$total_per_year_rent3."',
					'".$total_per_year_rent4."',
					'".$total_per_year_rent5."',
					'".$total_per_year_rent6."',
					'".$total."'
				)
			");
        }
        $numrow++;
    }
    header("location:index.php?numrow=$numrow");
?>
