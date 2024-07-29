package com.alibaba.easyexcel.test.temp.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

import com.alibaba.easyexcel.test.util.TestFileUtil;

import org.apache.easyexcel.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.easyexcel.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.easyexcel.poi.openxml4j.opc.OPCPackage;
import org.apache.easyexcel.poi.openxml4j.opc.PackageAccess;
import org.apache.easyexcel.poi.poifs.crypt.EncryptionInfo;
import org.apache.easyexcel.poi.poifs.crypt.EncryptionMode;
import org.apache.easyexcel.poi.poifs.crypt.Encryptor;
import org.apache.easyexcel.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.jupiter.api.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * 测试poi
 *
 * @author Jiaju Zhuang
 **/

public class Poi3Test {
    private static final Logger LOGGER = LoggerFactory.getLogger(Poi3Test.class);

    @Test
    public void Encryption() throws Exception {
        String file = TestFileUtil.getPath() + "large" + File.separator + "large07.xlsx";
        POIFSFileSystem fs = new POIFSFileSystem();
        EncryptionInfo info = new EncryptionInfo(EncryptionMode.agile);
        Encryptor enc = info.getEncryptor();
        enc.confirmPassword("foobaa");
        OPCPackage opc = OPCPackage.open(new File(file), PackageAccess.READ_WRITE);
        OutputStream os = enc.getDataStream(fs);
        opc.save(os);
        opc.close();

        // Write out the encrypted version
        FileOutputStream fos = new FileOutputStream("D:\\test\\99999999999.xlsx");
        fs.writeFilesystem(fos);
        fos.close();
        fs.close();

    }

    @Test
    public void Encryption2() throws Exception {
        Biff8EncryptionKey.setCurrentUserPassword("123456");
        POIFSFileSystem fs = new POIFSFileSystem(new File("d:/test/simple03.xls"), true);
        HSSFWorkbook hwb = new HSSFWorkbook(fs.getRoot(), true);
        Biff8EncryptionKey.setCurrentUserPassword(null);
        System.out.println(hwb.getSheetAt(0).getSheetName());

    }
}
