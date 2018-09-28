package com.rubixds;

import com.rubixds.reader.XSSFBalanceSheetReader;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.IOException;
import java.net.URISyntaxException;
import java.util.logging.Logger;

public class Driver {

    private static final Logger LOGGER = Logger.getLogger("Driver");

    public static void main(String[] args) {
        XSSFBalanceSheetReader balanceSheetReader = new XSSFBalanceSheetReader(3);
        try {
            balanceSheetReader.extractDataFromBalanceSheet();
        } catch (InvalidFormatException | IOException | URISyntaxException e) {
            e.printStackTrace();
        }

    }
}
