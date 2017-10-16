package com.alibaba.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Created by jipengfei on 17/3/19.
 * Field column num at excel head
 *
 * @author jipengfei
 * @date 2017/03/19
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelColumnNum {

    /**
     * col num
     * @return
     */
    int value();

    /**
     *
     * default @see com.alibaba.excel.util.TypeUtil
     * if default is not  meet you can set format
     *
     * @return
     */
    String format() default "";
}
