package com.xant.component;

public class TransactionContext {

    private static final String IN_TRANSACTION_KEY = "IN_TRANSACTION";

    public static void setInTransaction(boolean inTransaction) {
        ThreadContext.setAttribute(IN_TRANSACTION_KEY, inTransaction);
    }

    public static boolean getInTransaction() {
        return Boolean.TRUE.equals(ThreadContext.getAttribute(IN_TRANSACTION_KEY));
    }

    public static void clear() {
        ThreadContext.removeAttribute(IN_TRANSACTION_KEY);
    }
}
