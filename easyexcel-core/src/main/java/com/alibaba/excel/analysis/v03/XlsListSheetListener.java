package com.alibaba.excel.analysis.v03;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.easyexcel.poi.hssf.eventusermodel.EventWorkbookBuilder;
import org.apache.easyexcel.poi.hssf.eventusermodel.FormatTrackingHSSFListener;
import org.apache.easyexcel.poi.hssf.eventusermodel.HSSFEventFactory;
import org.apache.easyexcel.poi.hssf.eventusermodel.HSSFListener;
import org.apache.easyexcel.poi.hssf.eventusermodel.HSSFRequest;
import org.apache.easyexcel.poi.hssf.eventusermodel.MissingRecordAwareHSSFListener;
import org.apache.easyexcel.poi.hssf.record.BOFRecord;
import org.apache.easyexcel.poi.hssf.record.BoundSheetRecord;
import org.apache.easyexcel.poi.hssf.record.Record;

import com.alibaba.excel.analysis.v03.handlers.BofRecordHandler;
import com.alibaba.excel.analysis.v03.handlers.BoundSheetRecordHandler;
import com.alibaba.excel.context.xls.XlsReadContext;
import com.alibaba.excel.exception.ExcelAnalysisException;

/**
 * In some cases, you need to know the number of sheets in advance and only read the file once in advance.
 *
 * @author Jiaju Zhuang
 */
public class XlsListSheetListener implements HSSFListener {
    private final XlsReadContext xlsReadContext;
    private static final Map<Short, XlsRecordHandler> XLS_RECORD_HANDLER_MAP = new HashMap<Short, XlsRecordHandler>();

    static {
        XLS_RECORD_HANDLER_MAP.put(BOFRecord.sid, new BofRecordHandler());
        XLS_RECORD_HANDLER_MAP.put(BoundSheetRecord.sid, new BoundSheetRecordHandler());
    }

    public XlsListSheetListener(XlsReadContext xlsReadContext) {
        this.xlsReadContext = xlsReadContext;
        xlsReadContext.xlsReadWorkbookHolder().setNeedReadSheet(Boolean.FALSE);
    }

    @Override
    public void processRecord(Record record) {
        XlsRecordHandler handler = XLS_RECORD_HANDLER_MAP.get(record.getSid());
        if (handler == null) {
            return;
        }
        handler.processRecord(xlsReadContext, record);
    }

    public void execute() {
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        HSSFListener formatListener = new FormatTrackingHSSFListener(listener);
        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();
        EventWorkbookBuilder.SheetRecordCollectingListener workbookBuildingListener =
            new EventWorkbookBuilder.SheetRecordCollectingListener(formatListener);
        request.addListenerForAllRecords(workbookBuildingListener);
        try {
            factory.processWorkbookEvents(request, xlsReadContext.xlsReadWorkbookHolder().getPoifsFileSystem());
        } catch (IOException e) {
            throw new ExcelAnalysisException(e);
        }
    }
}
