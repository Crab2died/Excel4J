package utils;

import com.github.crab2died.utils.Utils;
import org.junit.Test;

import java.beans.IntrospectionException;
import java.math.BigDecimal;

public class UtilsTests {

    @Test
    public void getterAndSetter() {

        try {
            Utils.getterOrSetter(TestBean.class, "Afiled", Utils.FieldAccessType.GETTER);
            Utils.getterOrSetter(TestBean.class, "Afiled", Utils.FieldAccessType.SETTER);
            Utils.getterOrSetter(TestBean.class, "BFiled", Utils.FieldAccessType.GETTER);
            Utils.getterOrSetter(TestBean.class, "BFiled", Utils.FieldAccessType.SETTER);
        } catch (IntrospectionException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void bigDecimalTest() {
        BigDecimal bigDecimal = new BigDecimal("12989E19");
        System.out.println(bigDecimal);
    }
}
