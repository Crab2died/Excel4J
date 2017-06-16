package com.github.exceptions;

public class Excel4jWriteException extends Exception {

	private static final long serialVersionUID = -2186571940395162883L;

	public Excel4jWriteException() {
    }

    public Excel4jWriteException(String message) {
        super(message);
    }

    public Excel4jWriteException(String message, Throwable cause) {
        super(message, cause);
    }
}
