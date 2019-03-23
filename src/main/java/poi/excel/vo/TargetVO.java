package poi.excel.vo;

import java.util.List;

public class TargetVO {
    private List<?> list;

    public List<?> getList() {
        return this.list;
    }

    public void setList(List<?> list) {
        this.list = list;
    }

    public int setListCnt(List<?> list) {
        this.list = list;
        return list.size();
    }
}