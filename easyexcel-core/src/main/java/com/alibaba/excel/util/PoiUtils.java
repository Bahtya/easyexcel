package com.alibaba.excel.util;

import com.alibaba.excel.exception.ExcelRuntimeException;

import org.apache.easyexcel.poi.hssf.record.RowRecord;
import org.apache.easyexcel.poi.hssf.usermodel.HSSFRow;
import org.apache.easyexcel.poi.ss.usermodel.Row;
import org.apache.easyexcel.poi.util.BitField;
import org.apache.easyexcel.poi.util.BitFieldFactory;
import org.apache.easyexcel.poi.xssf.usermodel.XSSFRow;

import java.lang.reflect.Field;

/**
 * utils
 *
 * @author Jiaju Zhuang
 */
public class PoiUtils {

    /**
     * Whether to customize the height
     */
    public static final BitField CUSTOM_HEIGHT = BitFieldFactory.getInstance(0x640);

    private static final Field ROW_RECORD_FIELD = FieldUtils.getField(HSSFRow.class, "row", true);

    /**
     * Whether to customize the height
     *
     * @param row row
     * @return
     */
    public static boolean customHeight(Row row) {
        if (row instanceof XSSFRow) {
            XSSFRow xssfRow = (XSSFRow)row;
            return xssfRow.getCTRow().getCustomHeight();
        }
        if (row instanceof HSSFRow) {
            HSSFRow hssfRow = (HSSFRow)row;
            try {
                RowRecord record = (RowRecord)ROW_RECORD_FIELD.get(hssfRow);
                return CUSTOM_HEIGHT.getValue(record.getOptionFlags()) == 1;
            } catch (IllegalAccessException ignore) {
            }
        }
        return false;
    }
}
