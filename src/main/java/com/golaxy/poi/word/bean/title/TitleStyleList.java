package com.golaxy.poi.word.bean.title;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 标题层级列表
 * @author jiangzhaoyue
 */
public class TitleStyleList {
    private List<TitleStyle> list=new ArrayList<>();
    private Map<Integer,TitleStyle> map= new HashMap<>();

    /**
     * 设置默认级别
     */
    public TitleStyleList() {
        list.add(new TitleStyle(1,"标题1",18));
        list.add(new TitleStyle(2,"标题2",16));
        list.add(new TitleStyle(3,"标题3",14));
        list.add(new TitleStyle(4,"标题4",12));
        init();
    }

    /**
     * 自定义样式级别
     * @param list
     */
    public TitleStyleList(List<TitleStyle> list) {
        this.list = list;
        init();
    }
    private void init(){
        list.forEach(item->map.put(item.getLevel(),item));
    }

    /**
     * 根据级别返回样式名称
     * @param level
     * @return
     */
    public String getNameByLevel(Integer level){
        return map.get(level).getName();
    }

    public List<TitleStyle> getList() {
        return list;
    }
}
