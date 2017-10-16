package com.alibaba.excel.model;

import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

import com.alibaba.excel.annotation.FieldType;

/**
 * Created by jipengfei on 17/2/19.
 */
public class ReflectInfo {

    private Class<?> aClass;

    private String className;

    private Map<List<String>, FieldInfo> properties = new ConcurrentHashMap<List<String>, FieldInfo>();

    private Map<Integer, FieldInfo> properties1 = new ConcurrentHashMap<Integer, FieldInfo>();

    public void appendProperty(List<String> excelHead, String fieldName, FieldType type, String format) {
        properties.put(excelHead, new FieldInfo(fieldName, type, format));
    }

    public void appendProperty(Integer columnNum, String fieldName, FieldType type, String format) {
        properties1.put(columnNum, new FieldInfo(fieldName, type, format));
    }

    public FieldInfo getFieldName(List<String> excelHead) {
        return properties.get(excelHead);
    }

    public FieldInfo getFieldName(Integer column) {
        return properties1.get(column);
    }

    public Class<?> getaClass() {
        return aClass;
    }

    public void setaClass(Class<?> aClass) {
        this.aClass = aClass;
    }

    public String getClassName() {
        return className;
    }

    public void setClassName(String className) {
        this.className = className;
    }

    public class FieldInfo {

        private String fieldName;

        private FieldType type;

        private String format;

        public FieldInfo(String fieldName, FieldType type, String format) {
            this.fieldName = fieldName;
            this.type = type;
            this.format = format;

        }

        public String getFieldName() {
            return fieldName;
        }

        public void setFieldName(String fieldName) {
            this.fieldName = fieldName;
        }

        public FieldType getType() {
            return type;
        }

        public void setType(FieldType type) {
            this.type = type;
        }

        public String getFormat() {
            return format;
        }

        public void setFormat(String format) {
            this.format = format;
        }
    }

}
