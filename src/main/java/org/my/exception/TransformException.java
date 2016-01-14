package org.my.exception;

/**
 * 不支持的实体类字段类型
 * 
 * @author liu
 *
 */
public class TransformException extends RuntimeException {

	public TransformException() {
		super();
	}

	public TransformException(String message) {
		super(message);
	}

	public TransformException(String message, Throwable cause) {
		super(message, cause);
	}

	public TransformException(Throwable cause) {
		super(cause);
	}

	protected TransformException(String message, Throwable cause,
			boolean enableSuppression, boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

}
