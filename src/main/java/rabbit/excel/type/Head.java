package rabbit.excel.type;

import java.lang.annotation.*;

/**
 * 用于java bean字段上指定excel表头列名的注解
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Head {
    /**
     * 表头列名字，如果不指定，则默认为字段名
     *
     * @return 表头列名
     */
    String value() default "";
}

