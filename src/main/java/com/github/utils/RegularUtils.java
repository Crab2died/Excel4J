package com.github.utils;

import com.github.exceptions.IllegalGroupIndexException;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * <p>正则匹配相关工具</p>
 * author : Crab2Died</br>
 * date : 2017/5/24  9:43</br>
 */
public class RegularUtils {


    /**
     * <p>判断内容是否匹配</p></br>
     * author : Crab2Died</br>
     * date   : 2017年06月02日  15:46:25</br>
     *
     * @param pattern 匹配目标内容
     * @param reg     正则表达式
     * @return 返回boolean
     */
    static
    public boolean isMatched(String pattern, String reg) {
        Pattern compile = Pattern.compile(reg);
        return compile.matcher(pattern).matches();
    }

    /**
     * <p>正则提取匹配到的内容</p>
     * <p>例如：</p>
     * </br>
     * author : Crab2Died</br>
     * date   : 2017年06月02日  15:49:51</br>
     *
     * @param pattern 匹配目标内容
     * @param reg     正则表达式
     * @param group   提取内容索引
     * @return 提取内容集合
     * @throws {@link IllegalGroupIndexException}
     */
    static
    public List<String> match(String pattern, String reg, int group)
            throws IllegalGroupIndexException {

        List<String> matchGroups = new ArrayList<String>();
        Pattern compile = Pattern.compile(reg);
        Matcher matcher = compile.matcher(pattern);
        if (group > matcher.groupCount() || group < 0)
            throw new IllegalGroupIndexException("Illegal match group :" + group);
        while (matcher.find()) {
            matchGroups.add(matcher.group(group));
        }
        return matchGroups;
    }

    /**
     * <p>正则提取匹配到的内容,默认提取索引为0</p>
     * <p>例如：</p>
     * </br>
     * author : Crab2Died</br>
     * date   : 2017年06月02日  15:49:51</br>
     *
     * @param pattern 匹配目标内容
     * @param reg     正则表达式
     * @return 提取内容集合
     * @throws {@link IllegalGroupIndexException}
     */
    static public String match(String pattern, String reg) {

        String match = null;
        try {
            List<String> matchs = match(pattern, reg, 0);
            if (null != matchs && matchs.size() > 0) {
                match = matchs.get(0);
            }
        } catch (IllegalGroupIndexException e) {
            e.printStackTrace();
        }
        return match;
    }

    public static String converNumByReg(String number){
        Pattern compile = Pattern.compile("^(\\d+)(\\.0*)?$");
        Matcher matcher = compile.matcher(number);
        while (matcher.find()){
            number = matcher.group(1);
        }
        return number;
    }
}
