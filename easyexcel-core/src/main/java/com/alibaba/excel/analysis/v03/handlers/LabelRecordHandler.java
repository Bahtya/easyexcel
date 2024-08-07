package com.alibaba.excel.analysis.v03.handlers;

import com.alibaba.excel.analysis.v03.IgnorableXlsRecordHandler;
import com.alibaba.excel.context.xls.XlsReadContext;
import com.alibaba.excel.enums.RowTypeEnum;
import com.alibaba.excel.metadata.data.ReadCellData;

import org.apache.easyexcel.poi.hssf.record.LabelRecord;
import org.apache.easyexcel.poi.hssf.record.Record;

/**
 * Record handler
 *
 * @author Dan Zheng
 */
public class LabelRecordHandler extends AbstractXlsRecordHandler implements IgnorableXlsRecordHandler {
    @Override
    public void processRecord(XlsReadContext xlsReadContext, Record record) {
        LabelRecord lrec = (LabelRecord)record;
        String data = lrec.getValue();
        if (data != null && xlsReadContext.currentReadHolder().globalConfiguration().getAutoTrim()) {
            data = data.trim();
        }
        xlsReadContext.xlsReadSheetHolder().getCellMap().put((int)lrec.getColumn(),
            ReadCellData.newInstance(data, lrec.getRow(), (int)lrec.getColumn()));
        xlsReadContext.xlsReadSheetHolder().setTempRowType(RowTypeEnum.DATA);
    }
}
