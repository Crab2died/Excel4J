package com.github.crab2died.exceptions;

public class Excel4jReadException extends Exception {

	private static final long serialVersionUID = 8735084330744657672L;

	public Excel4jReadException() {

    }

    public Excel4jReadException(String message) {
        super(message);
    }

    public Excel4jReadException(String message, Throwable cause) {
        super(message, cause);
    }
}
