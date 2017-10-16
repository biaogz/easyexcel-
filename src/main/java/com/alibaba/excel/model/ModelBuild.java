package com.alibaba.excel.model;

import com.alibaba.excel.annotation.ExcelColumnNum;
import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.FieldType;
import com.alibaba.excel.model.ReflectInfo.FieldInfo;
import com.alibaba.excel.util.TypeUtil;

import org.apache.commons.beanutils.BeanUtils;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Created by jipengfei on 17/2/18.
 */
public class ModelBuild {

    private Map<Class<?>, ReflectInfo> cache = new ConcurrentHashMap<Class<?>, ReflectInfo>();

    public void buildHead(List<List<String>> head, Object objects) {

        List<String> strList = (List<String>)objects;
        for (int i = 0; i < strList.size(); i++) {
            List<String> list;
            if (head.size() <= i) {
                list = new ArrayList<String>();
                head.add(list);
            } else {
                list = head.get(0);
            }
            list.add(strList.get(i));
        }
    }

    public void buildModel(Object resultModel, Class<?> clazz, Object object, List<List<String>> head) {
        ReflectInfo reflectInfo = getReflectInfo(clazz);
        List<String> stringList = (List<String>)object;
        for (int i = 0; i < stringList.size(); i++) {
            if (head.size() > i) {
                List<String> oneColHead = head.get(i);
                FieldInfo fieldInfo = reflectInfo.getFieldName(i);
                if (fieldInfo == null) {
                    fieldInfo = reflectInfo.getFieldName(oneColHead);
                }
                try {
                    if (fieldInfo != null) {
                        Object value = TypeUtil.convert(stringList.get(i), fieldInfo.getType(), fieldInfo.getFormat());
                        if (value != null) {
                            BeanUtils.setProperty(resultModel, fieldInfo.getFieldName(), value);
                        }
                    } else {
                        System.out.println(oneColHead + "can not find in Class " + resultModel.getClass().getName());
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }

    }

    private ReflectInfo getReflectInfo(Class<?> clazz) {
        if (cache.get(clazz) == null) {
            ReflectInfo reflectInfo = new ReflectInfo();
            reflectInfo.setaClass(clazz);
            reflectInfo.setClassName(clazz.getName());
            Field[] fields = clazz.getDeclaredFields();
            for (int i = 0; i < fields.length; i++) {
                Field field = fields[i];
                String fieldName = field.getName();
                FieldType type = getFieldType(field.getType());
                ExcelProperty property = field.getAnnotation(ExcelProperty.class);
                if (property != null) {
                    String[] values = property.value();
                    String format = property.format();
                    List<String> excelHead = new ArrayList<String>();
                    excelHead.add(values[values.length - 1]);

                    reflectInfo.appendProperty(excelHead, fieldName, type, format);
                }
                ExcelColumnNum excelColumnNum = field.getAnnotation(ExcelColumnNum.class);
                if (excelColumnNum != null) {
                    Integer column = excelColumnNum.value();
                    String format = excelColumnNum.format();
                    reflectInfo.appendProperty(column, fieldName, type, format);
                }

            }
            cache.put(clazz, reflectInfo);
        }
        return cache.get(clazz);
    }

    private FieldType getFieldType(Class aclass) {
        if (String.class.equals(aclass)) {
            return FieldType.STRING;
        }
        if (Integer.class.equals(aclass) || int.class.equals(aclass)) {
            return FieldType.INT;
        }
        if (Double.class.equals(aclass) || double.class.equals(aclass)) {
            return FieldType.DOUBLE;
        }
        if (Boolean.class.equals(aclass) || boolean.class.equals(aclass)) {
            return FieldType.BOOLEAN;
        }
        if (Long.class.equals(aclass) || long.class.equals(aclass)) {
            return FieldType.LONG;
        }
        if (Date.class.equals(aclass)) {
            return FieldType.DATE;
        }
        return FieldType.STRING;

    }

}
