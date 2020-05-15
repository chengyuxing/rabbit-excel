package com.github.chengyuxing.excel.core;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Head {
    String value() default "";
}

