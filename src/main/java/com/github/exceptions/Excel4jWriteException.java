package com.github.exceptions;

public class Excel4jWriteException extends Exception {

    public Excel4jWriteException() {
    }

    public Excel4jWriteException(String message) {
        super(message);
    }

    public Excel4jWriteException(String message, Throwable cause) {
        super(message, cause);
    }
}
