package me.choicore.samples.excel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface Column {
    /**
     * 헤더
     *
     * @return header
     */
    String header();

    /**
     * -1 이면 자동으로 계산
     *
     * @return width
     */
    int width() default -1;

    String defaultValue() default "";

    HorizontalAlignment align() default HorizontalAlignment.LEFT;

    Style style() default Style.NORMAL;

    enum Style {
        /**
         * 기본 스타일
         */
        NORMAL,

        /**
         * 숫자 형식
         */
        NUMBER,

        /**
         * 날짜 형식 (날짜만)
         */
        DATE,

        /**
         * 통화 형식
         */
        CURRENCY,

        /**
         * 백분율 형식
         */
        PERCENTAGE

    }

}
