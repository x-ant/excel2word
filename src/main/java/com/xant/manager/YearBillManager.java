package com.xant.manager;

import com.xant.component.jdbc.plus.ServiceImpl;
import com.xant.dao.YearBillMapper;
import com.xant.entity.YearBillPO;

public class YearBillManager extends ServiceImpl<YearBillMapper, YearBillPO> {

    private YearBillManager() {
    }

    private static class SingletonHolder {
        private static final YearBillManager INSTANCE = new YearBillManager();
    }

    public static YearBillManager getInstance() {
        return SingletonHolder.INSTANCE;
    }

}
