<?php

namespace App\Http\Helper;

use Excel;
use PHPExcel_Exception;

ini_set('xdebug.var_display_max_depth', '-1');
ini_set('xdebug.var_display_max_children', '-1');
ini_set('xdebug.var_display_max_data', '-1');

class pexcel {

    /**
     * PHP Excel - bonus - version 1.0
     * use maatwebsite & phpoffice 
     * OPP
     * Author : Pham Phong
     * 
     * 
     */
    private $obj_PHPExcel;
    private $activated_sheet_index = 0;
    private $activated_sheet_name = 'new sheet';
    private $sheets_array = [];
    private $titles_array = [];
    private $values_array = [];
    //for multi area
    private $cells_formerge_array = [];

    //constuctor
    //destructor
    //Get
    //get sheets_array
    public function getSheets() {
        return $this->sheets_array;
    }

    //get titles_array
    public function getTitles() {
        return $this->titles_array;
    }

    //get values_array
    public function getValues() {
        return $this->values_array;
    }

    //get cells_merge_array
    public function getCellsMergeArray() {
        return $this->cells_formerge_array;
    }

    //get activated_sheet_index
    public function getActiveSheetIndex() {
        return $this->activated_sheet_index;
    }

    //SET
    //set active sheet index
    public function setActiveSheetIndex($activated_sheet_index) {
        $this->activated_sheet_index = $activated_sheet_index;
    }

    //set active sheet index
    public function setActiveSheetByName($activated_sheet_name) {
        $this->activated_sheet_name = $activated_sheet_name;
    }

    //set sheets array
    public function setSheetsArray($sheet_array) {
        $this->sheet_array = $sheet_array;
    }

    //set titles array
    public function setTitlesArray($title_array) {
        $this->titles_array = $title_array;
    }

    //set values array
    public function setValuesArray($value_array) {
        $this->values_array = $value_array;
    }

    //set cells merge array
    public function setCellsMergeArray($cells_formerging_array) {
        $this->cells_merge_array = cells_merge_array;
    }

    //import set object
    private function setOjbImport($path) {
        $this->obj_PHPExcel = Excel::load($path);
    }

    //export set object
    private function setObjExport($file_name) {
        $this->obj_PHPExcel = Excel::create($file_name);
    }

    //import - set a file into an array (const sheet)
    public function setExcelToArray($path) {
        $this->setOjbImport($path);

        $sheetData = $this->obj_PHPExcel
                ->getActiveSheet($this->activated_sheet_index)
                ->toArray(null, true, true, true);

        return $sheetData;
    }

    //set values for a sheet with value array
    public function setValuesForSheet($values_array, $begin_cell) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use($values_array, $begin_cell) {
                //fromArray($source, $nullValue, $startCell, $strictNullComparison, $headingGeneration)
                $sheet
                        ->fromArray($values_array, null, $begin_cell, false, true);
            });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->fromArray($values_array, null, $begin_cell, false, true);
        }
    }

    //set sheet styling
    public function setSheetStyling($style_array) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use($style_array) {
                $sheet->setStyle($style_array);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setStyle($style_array);
        }
    }

    //set sheet font
    public function setSheetFont($font_array) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use($font_array) {
                $sheet->setStyle($font_array);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setStyle($font_array);
        }
    }

    //set sheet fontsize
    public function setSheetFontSize($fontsize_int) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use($fontsize_int) {
                $sheet->setFontSize($fontsize_int);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setFontSize($fontsize_int);
        }
    }

    //set sheet font bold
    public function setSheetFontBold($fontbold_bool) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use($fontbold_bool) {
                $sheet->setFontBold($fontbold_bool);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setFontBold($fontbold_bool);
        }
    }

    //set sheet all border
    public function setSheetAllBorders($styleborder_string) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use($styleborder_string) {
                $sheet->setAllBorders($styleborder_string);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setAllBorders($styleborder_string);
        }
    }

    //set sheet border for a cell/ cells area   
    public function setSheetCellsBorder($cells_string, $styleborder_string) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use($cells_string, $styleborder_string) {
                $sheet->setBorder($cells_string, $styleborder_string);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setBorder($cells_string, $styleborder_string);
        }
    }

    //set value for a cell/ an area cells
    public function setValueForCell($cell_value, $cell_position_string) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use($cell_value, $cell_position_string) {
                $sheet->setCellValue($cell_position_string, $cell_value);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setCellValue($cell_position_string, $cell_value);
        }
    }

    //set auto filter for entire sheet
    public function setAutoFilter($area = false) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use ($area) {
                $sheet->setAutoFilter($area);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setAutoFilter($area);
        }
    }

    //set width for a column
    public function setColumnWidth($column_string, $width_int) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use ($column_string, $width_int) {
                $sheet->setWidth($column_string, $width_int);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setWidth($column_string, $width_int);
        }
    }

    //set width for columns
    public function setColumnsWidth($column_width_array) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use ($column_width_array) {
                $sheet->setWidth($column_width_array);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setWidth($column_width_array);
        }
    }

    //set height for a row
    public function setRowWidth($row_string, $height_int) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use ($row_string, $height_int) {
                $sheet->setHeight($row_string, $height_int);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setHeight($row_string, $height_int);
        }
    }

    //set height for rows
    public function setRowsWidth($row_height_array) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use ($row_height_array) {
                $sheet->setHeight($row_height_array);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setHeight($row_height_array);
        }
    }

    //set width+height for single cell
    public function setSizeSingle($cell_string, $width_int, $height_int) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use ($cell_string, $width_int, $height_int) {
                $sheet->setSize($cell_string, $width_int, $height_int);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setSize($cell_string, $width_int, $height_int);
        }
    }

    //set width+height for cell array
    public function setSizeArray($cell_array) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use ($cell_array) {
                $sheet->setSize($cell_array);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setSize($cell_array);
        }
    }

    //set enable/disable auto size for sheet
    public function setAutoSize($setting_bool) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use ($setting_bool) {
                $sheet->setAutoSize($setting_bool);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setAutoSize($setting_bool);
        }
    }

    //set disable auto size for column
    public function setDisableAutoSizeForColumn($Disable_Column_Bool_Array) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {
            $this->obj_PHPExcel->sheet($this->activated_sheet_name, function($sheet) use ($Disable_Column_Bool_Array) {
                $sheet->setAutoSize($Disable_Column_Bool_Array);
            });
        } else {
            $this->obj_PHPExcel
                    ->setActiveSheetIndex($this->activated_sheet_index)
                    ->setAutoSize($Disable_Column_Bool_Array);
        }
    }

    //set merging cells - single area
    public function setMergingCells($cells_formerge_string, $cell_value) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($cells_formerge_string, $cell_value) {
                        $sheet->mergeCells($cells_formerge_string);
                        $this->setValueForCell($cell_value, $cells_formerge_string);
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->mergeCells($cells_formerge_string);
        }
    }

    //set merging cells - multi area with array
    public function setMergingCellsArray($Merging_Cells_Array, $Value_Cells_Array) {
        foreach ($Merging_Cells_Array as $cells_formerge_string) {
            foreach ($Value_Cells_Array as $cell_value) {
                $this->setMergingCells($cells_formerge_string, $Value_Cells_Array);
                next($cells_formerge_string);
            }
        }
    }

    //set append ? row(s) after a row
    public function setAppendRow($after_row_int, $appended_array) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($after_row_int, $appended_array) {
                        $sheet->appendRow($after_row_int, $appended_array);
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->appendRow($after_row_int, $appended_array);
        }
    }

    //set prepended ? row(s) after a row
    public function setPrependedRow($before_row_int, $prependeded_array) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($before_row_int, $prependeded_array) {
                        $sheet->prependedRow($before_row_int, $prependeded_array);
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->prependedRow($before_row_int, $prependeded_array);
        }
    }

    //set background-color for a row
    public function setRowBgColor($row_int, $bg_color) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($row_int, $bg_color) {
                        $sheet->row($row_int, function($row) use ($bg_color) {
                            $row->setBackground($bg_color);
                        });
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->row($row, function($row) use ($bg_color) {
                        $row->setBackground($bg_color);
                    });
        }
    }

    //set value for a cell
    public function setCellValue($cell_string, $value) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($cell_string, $value) {
                        $sheet->cell($cell_string, function($cell) use ($value) {
                            $cell->setCellValue($value);
                        });
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->cell($cell_string, function($cell) use ($value) {
                        $cell->setCellValue($value);
                    });
        }
    }

    //set background-color for a cell
    public function setCellBgColor($cell_string, $bg_color) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($cell_string, $bg_color) {
                        $sheet->cell($cell_string, function($cell) use ($bg_color) {
                            $cell->setBackground($bg_color);
                        });
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->cell($cell_string, function($cell) use ($bg_color) {
                        $cell->setBackground($bg_color);
                    });
        }
    }

    //set background-color for cells area
    public function setCellsBgColor($cells_area_string, $bg_color) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($cells_area_string, $bg_color) {
                        $sheet->cells($cells_area_string, function($cells) use ($bg_color) {
                            $cells->setBackground($bg_color);
                        });
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->cells($cells_area_string, function($cells) use ($bg_color) {
                        $cells->setBackground($bg_color);
                    });
        }
    }

    //set font-color for cells area
    public function setCellsFontColor($cells_area_string, $font_color) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($cells_area_string, $font_color) {
                        $sheet->cells($cells_area_string, function($cells) use ($font_color) {
                            $cells->setFontColor($font_color);
                        });
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->cells($cells_area_string, function($cells) use ($font_color) {
                        $cells->setFontColor($font_color);
                    });
        }
    }

    //set font-family for cells area
    public function setCellsFontFamily($cells_area_string, $font_family) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($cells_area_string, $font_family) {
                        $sheet->cells($cells_area_string, function($cells) use ($font_family) {
                            $cells->setFontFamily($font_family);
                        });
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->cells($cells_area_string, function($cells) use ($font_family) {
                        $cells->setFontFamily($font_family);
                    });
        }
    }

    //set font-size for cells area
    public function setCellsFontSize($cells_area_string, $font_size) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($cells_area_string, $font_size) {
                        $sheet->cells($cells_area_string, function($cells) use ($font_size) {
                            $cells->setFontSize($font_size);
                        });
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->cells($cells_area_string, function($cells) use ($font_size) {
                        $cells->setFontSize($font_size);
                    });
        }
    }

    //set font-weight for cells area
    public function setCellsFontWeight($cells_area_string, $font_weight) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($cells_area_string, $font_weight) {
                        $sheet->cells($cells_area_string, function($cells) use ($font_weight) {
                            $cells->setFontWeight($font_weight);
                        });
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->cells($cells_area_string, function($cells) use ($font_weight) {
                        $cells->setFontWeight($font_weight);
                    });
        }
    }

    //set font array for cells area
    public function setCellsFont($cells_area_string, $font_array) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($cells_area_string, $font_array) {
                        $sheet->cells($cells_area_string, function($cells) use ($font_array) {
                            $cells->setFont($font_array);
                        });
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->cells($cells_area_string, function($cells) use ($font_array) {
                        $cells->setFont($font_array);
                    });
        }
    }

    //set font-border for cells area
    public function setCellsBorder($cells_area_string, $border_style_array = [
        'top' => 'solid',
        'right' => 'solid',
        'bottom' => 'solid',
        'left' => 'solid']) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($cells_area_string, $border_style_array) {
                        $sheet->cells($cells_area_string, function($cells) use ($border_style_array) {
                            $cells->setBorder($border_style_array);
                        });
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->cells($cells_area_string, function($cells) use ($border_style_array) {
                        $cells->setBorder($border_style_array);
                    });
        }
    }
    
    //set aligment for cells area
    public function setCellsAligment($cells_area_string, $aligment_style) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($cells_area_string, $aligment_style) {
                        $sheet->cells($cells_area_string, function($cells) use ($aligment_style) {
                            $cells->setAlignment($aligment_style);
                        });
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->cells($cells_area_string, function($cells) use ($aligment_style) {
                        $cells->setAlignment($aligment_style);
                    });
        }
    }
    
    //set Valigment for cells area
    public function setCellsValigment($cells_area_string, $valigment_style) {
        if ($this->obj_PHPExcel->getSheetByName($this->activated_sheet_name) == NULL) {

            $this->obj_PHPExcel
                    ->sheet($this->activated_sheet_name, function($sheet) use ($cells_area_string, $valigment_style) {
                        $sheet->cells($cells_area_string, function($cells) use ($valigment_style) {
                            $cells->setValignment($valigment_style);
                        });
                    });
        } else {
            $this->obj_PHPExcel->setActiveSheetIndex($this->activated_sheet_index)
                    ->cells($cells_area_string, function($cells) use ($valigment_style) {
                        $cells->setValignment($valigment_style);
                    });
        }
    }

    //print out
    //download
    public function excelDownload($file_type, $file_name = 'new_excel_file') {
        if (!empty($file_type) and ( $file_type == 'xls' or $file_type == 'xlsx' or $file_type == 'csv')) {
            $this->obj_PHPExcel->setFilename($file_name)->download($file_type);
        } else {
            throw new PHPExcel_Exception('We did not support this tail of excel : ' . $file_type);
        }
    }

}
