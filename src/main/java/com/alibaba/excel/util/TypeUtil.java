package com.alibaba.excel.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import com.alibaba.excel.annotation.FieldType;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;

/**
 * Created by jipengfei on 17/3/15.
 *
 * @author jipengfei
 * @date 2017/03/15
 */
public class TypeUtil {

    private static List<SimpleDateFormat> DATE_FORMAT_LIST = new ArrayList<SimpleDateFormat>(4);

    static {
        DATE_FORMAT_LIST.add(new SimpleDateFormat("yyyy/MM/dd HH:mm:ss"));
        DATE_FORMAT_LIST.add(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"));
    }

    public static Object convert(String value, FieldType type, String format) {
        if (isNotEmpty(value)) {
            try {

                if (FieldType.STRING.equals(type)) {
                    return value;
                }
                if (FieldType.INT.equals(type)) {
                    return Integer.parseInt(value);
                }
                if (FieldType.DOUBLE.equals(type)) {
                    return Double.parseDouble(value);
                }
                if (FieldType.BOOLEAN.equals(type)) {
                    String valueLower = value.toLowerCase();
                    if (valueLower.equals("true") || valueLower.equals("false")) {
                        return Boolean.parseBoolean(value.toLowerCase());
                    }
                    Integer integer = Integer.parseInt(value);
                    if (integer == 0) {
                        return false;
                    } else {
                        return true;
                    }
                }
                if (FieldType.LONG.equals(type)) {
                    return Long.parseLong(value);
                }
                if (FieldType.DATE.equals(type)) {
                    if (value.contains("-") || value.contains("/") || value.contains(":")) {
                        return getSimpleDateFormatDate(value, format);
                    } else {
                        Double d = Double.parseDouble(value);
                        return HSSFDateUtil.getJavaDate(d);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return null;
    }

    public static Date getSimpleDateFormatDate(String value, String format) {
        if (isNotEmpty(value)) {
            Date date = null;
            if (isNotEmpty(format)) {
                SimpleDateFormat simpleDateFormat = new SimpleDateFormat(format);
                try {
                    date = simpleDateFormat.parse(value);
                    return date;
                } catch (ParseException e) {
                }
            }
            for (SimpleDateFormat dateFormat : DATE_FORMAT_LIST) {
                try {
                    date = dateFormat.parse(value);
                } catch (ParseException e) {
                }
                if (date != null) {
                    break;
                }
            }

            return date;

        }
        return null;

    }

    private static Boolean isNotEmpty(String value) {
        if (value == null) {
            return false;
        }
        if (value.trim().equals("")) {
            return false;
        }
        return true;

    }

}
