package com.xant.manager;

import com.xant.component.jdbc.plus.ServiceImpl;
import com.xant.dao.YearBillMapper;
import com.xant.entity.YearBillPO;

import java.util.List;
import java.util.Map;

public class YearBillManager extends ServiceImpl<YearBillMapper, YearBillPO> {

    private YearBillManager() {
    }

    private static class SingletonHolder {
        private static final YearBillManager INSTANCE = new YearBillManager();
    }

    public static YearBillManager getInstance() {
        return SingletonHolder.INSTANCE;
    }

    public List<YearBillPO> queryListByCompany(String company) {
        return this.selectByMap(Map.of(YearBillPO.COMPANY, company));
    }

}
