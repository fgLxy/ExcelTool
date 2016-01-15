package org.my.exception;

public class UnSupportFileTypeException extends Exception {
	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public UnSupportFileTypeException() {
		super();
	}

	public UnSupportFileTypeException(String message) {
		super(message);
	}

	public UnSupportFileTypeException(String message, Throwable cause) {
		super(message, cause);
	}

	public UnSupportFileTypeException(Throwable cause) {
		super(cause);
	}

	protected UnSupportFileTypeException(String message, Throwable cause,
			boolean enableSuppression, boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}
}
