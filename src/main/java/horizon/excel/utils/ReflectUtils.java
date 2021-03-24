package horizon.excel.utils;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * @author K·J
 * <p>
 * Create at 2018-06-14 10:08
 */
public class ReflectUtils {
    /**
     * 获取成员变量的修饰符
     *
     * @param clazz class
     * @param field field
     * @param <T> object
     * @return T
     * @throws Exception error
     */
    public static <T> int getFieldModifier(Class<T> clazz, String field) throws Exception {
        Field[] fields = clazz.getDeclaredFields();//获取所有修饰符的成员变量，包括private,protected等getFields则不可以
        for (int i = 0; i < fields.length; i++) {
            if (fields[i].getName().equals(field)) {
                return fields[i].getModifiers();
            }
        }
        throw new Exception(clazz + " has no field [" + field + "]");
    }

    /**
     * 获取成员方法的修饰符
     *
     * @param clazz class
     * @param method method
     * @param <T> object
     * @return T
     * @throws Exception error
     */
    public static <T> int getMethodModifier(Class<T> clazz, String method) throws Exception {
        Method[] methods = clazz.getDeclaredMethods();
        for (int i = 0; i < methods.length; i++) {
            if (methods[i].getName().equals(method)) {
                return methods[i].getModifiers();
            }
        }
        throw new Exception(clazz + " has no method [" + method + "]");
    }

    /**
     * [对象]根据成员变量名称获取其值
     *
     * @param clazzInstance class instance
     * @param field field
     * @param <T> object
     * @return object
     * @throws NoSuchFieldException NoSuchFieldException
     * @throws SecurityException SecurityException
     * @throws IllegalArgumentException IllegalArgumentException
     * @throws IllegalAccessException IllegalAccessException
     * @throws NoSuchFieldException NoSuchFieldException
     */
    public static <T> Object getFieldValue(Object clazzInstance, Object field)
        throws SecurityException, IllegalArgumentException, IllegalAccessException ,NoSuchFieldException {
        Field[] fields = clazzInstance.getClass().getDeclaredFields();

        for (int i = 0; i < fields.length; i++) {
            if (fields[i].getName().equals(field)) {
                // 对于私有变量的访问权限，在这里设置，这样即可访问Private修饰的变量
                fields[i].setAccessible(true);
                return fields[i].get(clazzInstance);
            }
        }
        return null;
    }

    /**
     * [类]根据成员变量名称获取其值（默认值）
     *
     * @param clazz class
     * @param field field
     * @param <T> object
     * @return object
     * @throws NoSuchFieldException NoSuchFieldException
     * @throws SecurityException SecurityException
     * @throws IllegalArgumentException IllegalArgumentException
     * @throws IllegalAccessException IllegalAccessException
     * @throws InstantiationException InstantiationException
     */
    public static <T> Object getFieldValue(Class<T> clazz, String field)
        throws SecurityException, IllegalArgumentException, IllegalAccessException,
        InstantiationException,NoSuchFieldException {
        Field[] fields = clazz.getDeclaredFields();

        for (int i = 0; i < fields.length; i++) {
            if (fields[i].getName().equals(field)) {
                // 对于私有变量的访问权限，在这里设置，这样即可访问Private修饰的变量
                fields[i].setAccessible(true);
                return fields[i].get(clazz.newInstance());
            }
        }
        return null;
    }

    /**
     * 获取所有的成员变量(通过GET，SET方法获取)
     *
     * @param <T> object
     * @param clazz class
     * @return object
     */
    public static <T> String[] getFields(Class<T> clazz) {
        Field[] fields = clazz.getDeclaredFields();

        String[] fieldsArray = new String[fields.length];

        for (int i = 0; i < fields.length; i++) {
            fieldsArray[i] = fields[i].getName();
        }
        return fieldsArray;
    }

    /**
     * 获取所有的成员变量,包括父类
     *
     * @param <T> object
     * @param clazz class
     * @param superClass 是否包括父类
     * @return object
     */
    public static <T> Field[] getFields(Class<T> clazz, boolean superClass) {
        Field[] fields = clazz.getDeclaredFields();
        Field[] superFields = null;
        if (superClass) {
            Class superClazz = clazz.getSuperclass();
            if (superClazz != null) {
                superFields = superClazz.getDeclaredFields();
            }
        }

        Field[] allFields = null;

        if (superFields == null || superFields.length == 0) {
            allFields = fields;
        } else {
            allFields = new Field[fields.length + superFields.length];
            for (int i = 0; i < fields.length; i++) {
                allFields[i] = fields[i];
            }
            for (int i = 0; i < superFields.length; i++) {
                allFields[fields.length + i] = superFields[i];
            }
        }
        return allFields;
    }

    /**
     * 获取所有的成员变量,包括父类
     *
     * @param <T> object
     * @param clazz class
     * @return object
     * @throws Exception error
     */
    public static <T> Field[] getClassFieldsAndSuperClassFields(Class<T> clazz) throws Exception {
        Field[] fields = clazz.getDeclaredFields();

        if (clazz.getSuperclass() == null) {
            throw new Exception(clazz.getName() + "没有父类");
        }

        Field[] superFields = clazz.getSuperclass().getDeclaredFields();

        Field[] allFields = new Field[fields.length + superFields.length];

        for (int i = 0; i < fields.length; i++) {
            allFields[i] = fields[i];
        }
        for (int i = 0; i < superFields.length; i++) {
            allFields[fields.length + i] = superFields[i];
        }
        return allFields;
    }

    /**
     * 指定类，调用指定的无参方法
     *
     * @param <T> object
     * @param clazz class
     * @param method method
     * @throws NoSuchMethodException NoSuchMethodException
     * @throws SecurityException SecurityException
     * @throws IllegalAccessException IllegalAccessException
     * @throws IllegalArgumentException IllegalArgumentException
     * @throws InvocationTargetException InvocationTargetException
     * @throws InstantiationException InstantiationException
     * @return object
     */
    public static <T> Object invoke(Class<T> clazz, String method)
        throws NoSuchMethodException, SecurityException, IllegalAccessException,
        IllegalArgumentException, InvocationTargetException, InstantiationException {
        Object instance = clazz.newInstance();
        Method m = clazz.getMethod(method, new Class[]{});
        return m.invoke(instance, new Object[]{});
    }

    /**
     * 通过对象，访问其方法
     *
     * @param <T> object
     * @param clazzInstance class instance
     * @param method method
     * @return object
     * @throws NoSuchMethodException NoSuchMethodException
     * @throws SecurityException SecurityException
     * @throws IllegalAccessException IllegalAccessException
     * @throws IllegalArgumentException IllegalArgumentException
     * @throws InvocationTargetException InvocationTargetException
     */
    public static <T> Object invoke(Object clazzInstance, String method)
        throws NoSuchMethodException, SecurityException, IllegalAccessException,
        IllegalArgumentException, InvocationTargetException {
        Method m = clazzInstance.getClass().getMethod(method, new Class[]{});
        return m.invoke(clazzInstance, new Object[]{});
    }

    /**
     * 指定类，调用指定的方法
     *
     * @param <T> object
     * @param clazz class
     * @param method method
     * @param paramClasses param class
     * @param params param
     * @return Object
     * @throws InstantiationException InstantiationException
     * @throws IllegalAccessException IllegalAccessException
     * @throws NoSuchMethodException NoSuchMethodException
     * @throws SecurityException SecurityException
     * @throws IllegalArgumentException IllegalArgumentException
     * @throws InvocationTargetException InvocationTargetException
     */
    public static <T> Object invoke(Class<T> clazz, String method,
                                    Class<T>[] paramClasses, Object[] params)
        throws InstantiationException, IllegalAccessException,
        NoSuchMethodException, SecurityException, IllegalArgumentException, InvocationTargetException {
        Object instance = clazz.newInstance();
        Method _m = clazz.getMethod(method, paramClasses);
        return _m.invoke(instance, params);
    }

    /**
     * 通过类的实例，调用指定的方法
     *
     * @param <T> object
     * @param clazzInstance class instance
     * @param method method
     * @param paramClasses param classes
     * @param params param
     * @return object
     * @throws IllegalAccessException IllegalAccessException
     * @throws NoSuchMethodException NoSuchMethodException
     * @throws SecurityException SecurityException
     * @throws IllegalArgumentException IllegalArgumentException
     * @throws InvocationTargetException InvocationTargetException
     */
    public static <T> Object invoke(Object clazzInstance, String method,
                                    Class<T>[] paramClasses, Object[] params)
        throws IllegalAccessException, NoSuchMethodException,
        SecurityException, IllegalArgumentException, InvocationTargetException {
        Method _m = clazzInstance.getClass().getMethod(method, paramClasses);
        return _m.invoke(clazzInstance, params);
    }
}
