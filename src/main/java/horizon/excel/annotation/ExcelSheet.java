package horizon.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author K·J
 * <p>
 * Create at 2018-06-14 9:58
 */
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD, ElementType.TYPE})
public @interface ExcelSheet {
    //sheet名
    String name() default "";

    //导入时序号
    int importIndex() default 0;

    //导出时序号
    int exportIndex() default 0;
}
