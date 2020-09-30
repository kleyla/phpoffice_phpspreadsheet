 public function excel($id)
    {
        // dd('hi');
        $presupuesto = Presupuesto::find($id);
        $spreadsheet = new Spreadsheet();
        $sheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, 'Formulario');
        $spreadsheet->addSheet($sheet, 0);
        $sheet->setCellValue('A1', 'FORMULARIO DE PROPUESTA');
        $sheet->getStyle('A1')->getAlignment()->setHorizontal('center');
        $sheet->mergeCells('A1:F1');
        $sheet->setCellValue('A2', $presupuesto->codigo);
        $nCellIni = 3;
        $sheet->setCellValue('A3', 'ITEM');
        $sheet->setCellValue('B3', 'DESCRIPCIÓN ITEM');
        $sheet->setCellValue('C3', 'UNIDAD');
        $sheet->setCellValue('D3', 'CANTIDAD');
        $sheet->setCellValue('E3', 'PARCIAL');
        $sheet->setCellValue('F3', 'TOTAL');
        $sheet->getStyle('A3:F3')->getAlignment()->setHorizontal('center');

        $styleBold = [
            'font' => [
                'bold' => true,
            ],
        ];
        $styleBorder = [
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ],
            ],
        ];
        $styleBorderVertical = [
            'borders' => [
                'vertical' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ],
                'outline' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ],
            ],
        ];
        $sheet->getStyle('A1:F3')->applyFromArray($styleBold);

        $items = DB::select("Select mp_item.nombre, mp_item.unidad, p_items.cantidad,
                    p_items.costo, p_items.total
                from p_items, mp_item
                where p_items.id_item = mp_item.id and p_items.id_presupuesto = $id");
        $nCell = 4;
        $nro = 1;
        foreach ($items as $item) {
            $sheet->setCellValue('A' . $nCell, $nro);
            $sheet->setCellValue('B' . $nCell, $item->nombre);
            $sheet->setCellValue('C' . $nCell, $item->unidad);
            $sheet->setCellValue('D' . $nCell, $item->cantidad);
            $sheet->setCellValue('E' . $nCell, $item->costo);
            $sheet->setCellValue('F' . $nCell, $item->total);
            $nCell = $nCell + 1;
            $nro = $nro + 1;
        }
        $sheet->getStyle('A4:A' . $nCell)->getAlignment()->setHorizontal('center');
        $sheet->getStyle('C4:C' . $nCell)->getAlignment()->setHorizontal('center');
        $sheet->setCellValue('A' . $nCell, 'TOTAL (Bs)');
        $sheet->setCellValue('F' . $nCell, $presupuesto->total);
        $sheet->getStyle('A' . $nCell . ':F' . $nCell)->applyFromArray($styleBold);
        $sheet->getStyle('A' . $nCellIni . ':F' . $nCell)->applyFromArray($styleBorder);
        // CENTER
        $sheet->getColumnDimension('A')->setAutoSize(true);
        $sheet->getColumnDimension('B')->setAutoSize(true);
        $sheet->getColumnDimension('C')->setAutoSize(true);
        $sheet->getColumnDimension('D')->setAutoSize(true);
        $sheet->getColumnDimension('E')->setAutoSize(true);
        $sheet->getColumnDimension('F')->setAutoSize(true);

        // NEW PAGE
        $sheetPU = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($spreadsheet, 'Analisis PU');
        $spreadsheet->addSheet($sheetPU, 1);
        $sheetPU->setCellValue('A1', 'ANÁLISIS');
        $sheetPU->getStyle('A1')->getAlignment()->setHorizontal('center');
        $sheetPU->mergeCells('A1:G1');

        $cliente = Cliente::find($presupuesto->id_cliente);
        $sheetPU->setCellValue('A3', 'OBRA');
        $sheetPU->setCellValue('B3', $presupuesto->codigo);
        $sheetPU->setCellValue('A4', 'EMPRESA');
        $sheetPU->setCellValue('B4', $cliente->nombre);
        $sheetPU->setCellValue('C4', 'DIRECCIÓN');
        $sheetPU->setCellValue('D4', $cliente->direccion);
        $sheetPU->getStyle('A1:F4')->applyFromArray($styleBold);

        $datos = CostosIndirectos::find($presupuesto->id_costo_i);
        $t_items =  DB::select("Select p_items.id_i, p_items.cantidad, p_items.costo, p_items.total, mp_item.nombre, mp_item.unidad
                from p_items, mp_item
                where p_items.id_item = mp_item.id and p_items.id_presupuesto = $id");
        // dd($t_items);
        $nroItem = 1;
        $nCell = 5;
        $nro = 1;
        for ($i = 0; $i < sizeof($t_items); $i++) {
            $sheetPU->setCellValue('B' . $nCell, 'Item');
            $sheetPU->setCellValue('C' . $nCell, $nro . ' ' . $t_items[$i]->nombre);
            $nro = $nro + 1;
            $sheetPU->setCellValue('D' . $nCell, 'Unidad');
            $sheetPU->setCellValue('E' . $nCell, $t_items[$i]->unidad);
            $sheetPU->getStyle('B' . $nCell . ':G' . $nCell)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('bcd4e6');
            $nCell = $nCell + 1;
            $nCellIni = $nCell;
            $sheetPU->setCellValue('B' . $nCell, 'No');
            $sheetPU->getStyle('B' . $nCell)->getAlignment()->setVertical('center');
            $sheetPU->mergeCells('B' . $nCell . ':B' . ($nCell + 1));
            $sheetPU->setCellValue('C' . $nCell, 'DESCRIPCIÓN');
            $sheetPU->getStyle('C' . $nCell)->getAlignment()->setVertical('center');
            $sheetPU->mergeCells('C' . $nCell . ':C' . ($nCell + 1));
            $sheetPU->setCellValue('D' . $nCell, 'UNIDAD');
            $sheetPU->getStyle('D' . $nCell)->getAlignment()->setVertical('center');
            $sheetPU->mergeCells('D' . $nCell . ':D' . ($nCell + 1));
            $sheetPU->setCellValue('E' . $nCell, 'RENDIMIENTO');
            $sheetPU->getStyle('E' . $nCell)->getAlignment()->setVertical('center');
            $sheetPU->mergeCells('E' . $nCell . ':E' . ($nCell + 1));
            $sheetPU->setCellValue('F' . $nCell, 'PRECIOS (BS.)');
            $sheetPU->mergeCells('F' . $nCell . ':G' . $nCell);
            $sheetPU->getStyle('A' . $nCell . ':G' . $nCell)->applyFromArray($styleBold);
            $nCell = $nCell + 1;
            $sheetPU->setCellValue('F' . $nCell, 'UNITARIO');
            $sheetPU->setCellValue('G' . $nCell, 'PARCIAL');
            $sheetPU->getStyle('A' . $nCell . ':G' . $nCell)->applyFromArray($styleBold);
            $sheetPU->getStyle('B' . $nCellIni . ':G' . $nCell)->applyFromArray($styleBorder);
            $sheetPU->getStyle('B' . $nCellIni . ':G' . $nCell)->getAlignment()->setHorizontal('center');
            $nCell = $nCell + 1;
            $nCellIni = $nCell;
            $sheetPU->setCellValue('B' . $nCell, 'A');
            $sheetPU->setCellValue('C' . $nCell, 'MATERIALES');
            $sheetPU->getStyle('A' . $nCell . ':G' . $nCell)->applyFromArray($styleBold);
            $nCell = $nCell + 1;
            // $id = $t_items[$i]->id_i;
            // MATERIALES
            $mat = 0;
            $nroInsumo = 1;
            $insumos = $this->get_insumos($t_items[$i]->id_i, 'Materiales');
            for ($e = 0; $e < sizeof($insumos); $e++) {
                $mat = $mat + floatval($insumos[$e]->total);
                $sheetPU->setCellValue('B' . $nCell, $nroInsumo);
                $sheetPU->setCellValue('C' . $nCell, $insumos[$e]->nombre);
                $sheetPU->setCellValue('D' . $nCell, $insumos[$e]->unidad);
                $sheetPU->setCellValue('E' . $nCell, $insumos[$e]->cantidad);
                $sheetPU->setCellValue('F' . $nCell, $insumos[$e]->costo);
                $sheetPU->setCellValue('G' . $nCell, $insumos[$e]->total);
                $nCell = $nCell + 1;
                $nroInsumo = $nroInsumo + 1;
            }
            $sheetPU->getStyle('B' . $nCellIni . ':G' . ($nCell - 1))->applyFromArray($styleBorderVertical);
            $sheetPU->setCellValue('F' . $nCell, 'SUBTOTAL (A)');
            $sheetPU->setCellValue('G' . $nCell, $mat);
            // MANO DE OBRA
            $mo = 0;
            $insumos = $this->get_insumos($t_items[$i]->id_i, 'Mano de Obra');
            $nCell = $nCell + 2; //espacio
            $nCellIni = $nCell;
            $sheetPU->setCellValue('B' . $nCell, 'B');
            $sheetPU->setCellValue('C' . $nCell, 'MANO DE OBRA');
            $sheetPU->getStyle('A' . $nCell . ':G' . $nCell)->applyFromArray($styleBold);
            $nCell = $nCell + 1;
            for ($e = 0; $e < sizeof($insumos); $e++) {
                $mo = $mo + floatval($insumos[$e]->total);
                $sheetPU->setCellValue('B' . $nCell, $nroInsumo);
                $sheetPU->setCellValue('C' . $nCell, $insumos[$e]->nombre);
                $sheetPU->setCellValue('D' . $nCell, $insumos[$e]->unidad);
                $sheetPU->setCellValue('E' . $nCell, $insumos[$e]->cantidad);
                $sheetPU->setCellValue('F' . $nCell, $insumos[$e]->costo);
                $sheetPU->setCellValue('G' . $nCell, $insumos[$e]->total);
                $nCell = $nCell + 1;
                $nroInsumo = $nroInsumo + 1;
            }
            $sheetPU->getStyle('B' . $nCellIni . ':G' . ($nCell - 1))->applyFromArray($styleBorderVertical);
            $sheetPU->setCellValue('F' . $nCell, 'SUBTOTAL (B)');
            $sheetPU->setCellValue('G' . $nCell, $mo);

            // EQUIPO Y MAQ
            $eq_maq = 0;
            $insumos = $this->get_insumos($t_items[$i]->id_i, 'Equipo y maquinaria');
            $nCell = $nCell + 2; //espacio
            $nCellIni = $nCell;
            $sheetPU->setCellValue('B' . $nCell, 'C');
            $sheetPU->setCellValue('C' . $nCell, 'HERRAMIENTAS Y EQUIPOS');
            $sheetPU->getStyle('A' . $nCell . ':G' . $nCell)->applyFromArray($styleBold);
            $nCell = $nCell + 1;
            for ($e = 0; $e < sizeof($insumos); $e++) {
                $eq_maq = $eq_maq + floatval($insumos[$e]->total);
                $sheetPU->setCellValue('B' . $nCell, $nroInsumo);
                $sheetPU->setCellValue('C' . $nCell, $insumos[$e]->nombre);
                $sheetPU->setCellValue('D' . $nCell, $insumos[$e]->unidad);
                $sheetPU->setCellValue('E' . $nCell, $insumos[$e]->cantidad);
                $sheetPU->setCellValue('F' . $nCell, $insumos[$e]->costo);
                $sheetPU->setCellValue('G' . $nCell, $insumos[$e]->total);
                $nCell = $nCell + 1;
                $nroInsumo = $nroInsumo + 1;
            }
            $sheetPU->getStyle('B' . $nCellIni . ':G' . ($nCell - 1))->applyFromArray($styleBorderVertical);
            $sheetPU->setCellValue('F' . $nCell, 'SUBTOTAL (C)');
            $sheetPU->setCellValue('G' . $nCell, $eq_maq);
            $nCell = $nCell + 2;
            $nCellIni = $nCell;
            $sheetPU->setCellValue('B' . $nCell, 'D');
            $sheetPU->setCellValue('C' . $nCell, 'TOTAL COSTO DIRECTO - SUBTOTAL (A+B+C)');
            $sheetPU->getStyle('C' . $nCell)->getAlignment()->setHorizontal('right');
            $sheetPU->getStyle('B' . $nCell . ':G' . $nCell)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('cccccc');
            $sheetPU->mergeCells('C' . $nCell . ':E' . $nCell);
            $sub = $mat + $mo + $eq_maq;
            $sheetPU->setCellValue('G' . $nCell, $sub);
            $nCell = $nCell + 1;
            // CARGA SOCIAL
            $cs = $mo * (floatval($datos->cargas_sociales) / 100);
            $cs = round($cs, 2);
            $sheetPU->setCellValue('B' . $nCell, 'E');
            $sheetPU->setCellValue('C' . $nCell, 'BENIFICIOS SOCIALES MANO DE OBRA ' . $datos->cargas_sociales . ' % de(B)');
            $sheetPU->getStyle('C' . $nCell)->getAlignment()->setHorizontal('right');
            $sheetPU->mergeCells('C' . $nCell . ':E' . $nCell);
            $sheetPU->setCellValue('G' . $nCell, $cs);
            $nCell = $nCell + 1;
            // IVA
            $iva = ($mo + $cs) * (floatval($datos->iva_mo_cs) / 100);
            $iva = round($iva, 2);
            $sheetPU->setCellValue('B' . $nCell, 'F');
            $sheetPU->setCellValue('C' . $nCell, 'I.V.A. de M.O. y cargas sociales ' . $datos->iva_mo_cs . ' % de(B y E)');
            $sheetPU->getStyle('C' . $nCell)->getAlignment()->setHorizontal('right');
            $sheetPU->mergeCells('C' . $nCell . ':E' . $nCell);
            $sheetPU->setCellValue('G' . $nCell, $iva);
            $nCell = $nCell + 1;
            // HERRAMIENTAS
            $total_mo = $mo + $cs + $iva;
            $herr = $total_mo  * (floatval($datos->herramientas) / 100);
            $herr = round($herr, 2);
            $sheetPU->setCellValue('B' . $nCell, 'G');
            $sheetPU->setCellValue('C' . $nCell, 'HERRAMIENTAS ' . $datos->herramientas . ' % de(B, E y F)');
            $sheetPU->getStyle('C' . $nCell)->getAlignment()->setHorizontal('right');
            $sheetPU->mergeCells('C' . $nCell . ':E' . $nCell);
            $sheetPU->setCellValue('G' . $nCell, $herr);
            $nCell = $nCell + 1;
            $total_eq_maq = $eq_maq + $herr;

            // GG
            $gg = ($mat + $total_mo + $total_eq_maq) * (floatval($datos->gastos_generales) / 100);
            $gg = round($gg, 2);
            $sheetPU->setCellValue('B' . $nCell, 'H');
            $sheetPU->setCellValue('C' . $nCell, 'GASTOS GENERALES Y ADMINISTRATIVOS ' . $datos->gastos_generales . ' % de(A+B+C+D+E+F+G)');
            $sheetPU->getStyle('C' . $nCell)->getAlignment()->setHorizontal('right');
            $sheetPU->mergeCells('C' . $nCell . ':E' . $nCell);
            $sheetPU->setCellValue('G' . $nCell, $gg);
            $nCell = $nCell + 1;
            // UTILIDAD
            $utilidad = ($mat + $total_mo + $total_eq_maq + $gg) * (floatval($datos->utilidad) / 100);
            $utilidad = round($utilidad, 2);
            $sheetPU->setCellValue('B' . $nCell, 'I');
            $sheetPU->setCellValue('C' . $nCell, 'UTILIDADES ' . $datos->utilidad . ' % de(A+B+C+D+E+F+G+H)');
            $sheetPU->getStyle('C' . $nCell)->getAlignment()->setHorizontal('right');
            $sheetPU->mergeCells('C' . $nCell . ':E' . $nCell);
            $sheetPU->setCellValue('G' . $nCell, $utilidad);
            $nCell = $nCell + 1;
            // IT
            $it = ($mat + $total_mo + $total_eq_maq + $gg + $utilidad) * (floatval($datos->it) / 100);
            $it = round($it, 2);
            $sheetPU->setCellValue('B' . $nCell, 'J');
            $sheetPU->setCellValue('C' . $nCell, 'IT ' . $datos->it . ' % de(A+B+C+D+E+F+G+H+I)');
            $sheetPU->getStyle('C' . $nCell)->getAlignment()->setHorizontal('right');
            $sheetPU->mergeCells('C' . $nCell . ':E' . $nCell);
            $sheetPU->setCellValue('G' . $nCell, $it);
            $nCell = $nCell + 1;
            $ci = $cs + $iva + $herr + $gg + $utilidad + $it;
            $ci = round($ci, 2);
            $sheetPU->setCellValue('B' . $nCell, 'K');
            $sheetPU->setCellValue('C' . $nCell, 'TOTAL COSTO INDIRECTO - SUBTOTAL (E+F+G+H+I+J)');
            $sheetPU->getStyle('C' . $nCell)->getAlignment()->setHorizontal('right');
            $sheetPU->getStyle('B' . $nCell . ':G' . $nCell)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('cccccc');
            $sheetPU->mergeCells('C' . $nCell . ':E' . $nCell);
            $sheetPU->setCellValue('G' . $nCell, $ci);
            $nCell = $nCell + 1;
            $total = $sub + $ci;
            $total = round($total, 2);
            $sheetPU->setCellValue('B' . $nCell, 'L');
            $sheetPU->setCellValue('C' . $nCell, 'TOTAL (D+K)');
            $sheetPU->getStyle('C' . $nCell)->getAlignment()->setHorizontal('right');
            $sheetPU->mergeCells('C' . $nCell . ':E' . $nCell);
            $sheetPU->setCellValue('G' . $nCell, $total);
            $nCell = $nCell + 1;
            $sheetPU->setCellValue('B' . $nCell, 'M');
            $sheetPU->setCellValue('C' . $nCell, 'PRECIO UNITARIO ADOPTADO');
            $sheetPU->getStyle('C' . $nCell)->getAlignment()->setHorizontal('center');
            $sheetPU->getStyle('B' . $nCell . ':G' . $nCell)->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)->getStartColor()->setARGB('e3dac9');
            $sheetPU->mergeCells('C' . $nCell . ':E' . $nCell);
            $sheetPU->setCellValue('G' . $nCell, $total);
            $sheetPU->getStyle('B' . $nCellIni . ':G' . $nCell)->applyFromArray($styleBorderVertical);

            $nCell = $nCell + 5;
            $sheetPU->setCellValue('C' . $nCell, 'REPRESENTANTE LEGAL');
            $sheetPU->getStyle('C' . $nCell)->getAlignment()->setHorizontal('center');

            $nCell = $nCell + 2; //espacio
        }
        // CENTER
        $sheetPU->getStyle('B7:B' . $nCell)->getAlignment()->setHorizontal('center');
        $sheetPU->getStyle('D7:D' . $nCell)->getAlignment()->setHorizontal('center');
        // WITH OF COLUMNS
        $sheetPU->getColumnDimension('B')->setAutoSize(true);
        $sheetPU->getColumnDimension('C')->setAutoSize(true);
        $sheetPU->getColumnDimension('D')->setAutoSize(true);
        $sheetPU->getColumnDimension('E')->setAutoSize(true);
        $sheetPU->getColumnDimension('F')->setAutoSize(true);
        $sheetPU->getColumnDimension('G')->setAutoSize(true);

        $date = now();
        $filename = "presupuesto $date.xlsx";

        try {
            $writer = new Xlsx($spreadsheet);
            $writer->save($filename);
            $content = file_get_contents($filename);
        } catch (Exception $e) {
            exit($e->getMessage());
        }

        header("Content-Disposition: attachment; filename=" . $filename);
        unlink($filename);
        exit($content);
    }
