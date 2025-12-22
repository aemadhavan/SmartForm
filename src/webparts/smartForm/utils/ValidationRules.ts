export class ValidationRules {
  /**
   * Validates if a string is a valid email address
   */
  public static isValidEmail(email: string): boolean {
    if (!email || typeof email !== 'string') {
      return false;
    }
    // RFC 5322 simplified email regex
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email.trim());
  }

  /**
   * Validates if a string is a valid URL
   */
  public static isValidUrl(url: string): boolean {
    if (!url || typeof url !== 'string') {
      return false;
    }
    try {
      const urlObj = new URL(url);
      // Only allow http and https protocols for security
      return urlObj.protocol === 'http:' || urlObj.protocol === 'https:';
    } catch {
      return false;
    }
  }

  /**
   * Validates if a value is a valid date
   */
  public static isValidDate(date: any): boolean {
    if (!date) {
      return false;
    }
    if (date instanceof Date) {
      return !isNaN(date.getTime());
    }
    if (typeof date === 'string') {
      const parsedDate = new Date(date);
      return !isNaN(parsedDate.getTime());
    }
    return false;
  }

  /**
   * Validates if a number is within a specified range
   */
  public static isWithinRange(
    value: number,
    min?: number,
    max?: number
  ): boolean {
    if (typeof value !== 'number' || isNaN(value)) {
      return false;
    }
    if (min !== undefined && value < min) {
      return false;
    }
    if (max !== undefined && value > max) {
      return false;
    }
    return true;
  }

  /**
   * Validates if a string length is within specified range
   */
  public static isValidLength(
    text: string,
    minLength?: number,
    maxLength?: number
  ): boolean {
    if (typeof text !== 'string') {
      return false;
    }
    const length = text.length;
    if (minLength !== undefined && length < minLength) {
      return false;
    }
    if (maxLength !== undefined && length > maxLength) {
      return false;
    }
    return true;
  }

  /**
   * Validates if a date is within a specified range
   */
  public static isDateInRange(
    date: Date,
    minDate?: string,
    maxDate?: string
  ): boolean {
    if (!this.isValidDate(date)) {
      return false;
    }
    const dateValue = date.getTime();

    if (minDate) {
      const minDateTime = new Date(minDate).getTime();
      if (dateValue < minDateTime) {
        return false;
      }
    }

    if (maxDate) {
      const maxDateTime = new Date(maxDate).getTime();
      if (dateValue > maxDateTime) {
        return false;
      }
    }

    return true;
  }

  /**
   * Sanitizes a string for safe use in file names
   */
  public static sanitizeFileName(fileName: string): string {
    // Remove or replace characters that are invalid in file names
    return fileName
      .replace(/[<>:"/\\|?*]/g, '_')
      .replace(/\s+/g, '_')
      .replace(/_{2,}/g, '_');
  }
}
