package com.github.exceptions;

public class Excel4jReadException extends Exception {

    public Excel4jReadException() {

    }

    public Excel4jReadException(String message) {
        super(message);
    }

    public Excel4jReadException(String message, Throwable cause) {
        super(message, cause);
    }
}
