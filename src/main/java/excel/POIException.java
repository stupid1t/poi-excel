package excel;

import java.security.InvalidParameterException;

/**
 * 625
 * 
 * @author 625
 *
 */
public class POIException extends RuntimeException {

	private String errCode;

	public String getErrCode() {
		return errCode;
	}

	public POIException setErrCode(String errCode) {
		this.errCode = errCode;
		return this;
	}

	private static final long serialVersionUID = -2282295769920642919L;

	public POIException(String message) {
		super(message);
	}

	public static POIException newMessageException(String message) {
		return new POIException(message);
	}

	public static POIException newMessageException(String message, Object... args) {
		try {
			for (Object arg : args) {
				message = message.replace("{}", String.valueOf(arg));
			}
		} catch (Exception e) {
			throw new InvalidParameterException();
		}
		return new POIException(message);
	}

}
