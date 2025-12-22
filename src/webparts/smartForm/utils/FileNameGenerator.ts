import { ValidationRules } from './ValidationRules';

export class FileNameGenerator {
  /**
   * Generates a file name based on a pattern with placeholders
   * Supported placeholders:
   * - {timestamp} - Unix timestamp
   * - {userName} - Sanitized user display name
   * - {date} - Current date in YYYY-MM-DD format
   * - {time} - Current time in HH-MM-SS format
   * - {guid} - Random GUID
   */
  public static generateFileName(pattern: string, userName: string): string {
    const now = new Date();

    let fileName = pattern;

    // Replace timestamp
    fileName = fileName.replace(/{timestamp}/g, now.getTime().toString());

    // Replace userName (sanitized)
    const sanitizedUserName = ValidationRules.sanitizeFileName(userName);
    fileName = fileName.replace(/{userName}/g, sanitizedUserName);

    // Replace date (YYYY-MM-DD)
    const dateStr = this.formatDate(now);
    fileName = fileName.replace(/{date}/g, dateStr);

    // Replace time (HH-MM-SS)
    const timeStr = this.formatTime(now);
    fileName = fileName.replace(/{time}/g, timeStr);

    // Replace guid
    fileName = fileName.replace(/{guid}/g, this.generateGuid());

    // Ensure .json extension
    if (!fileName.toLowerCase().endsWith('.json')) {
      fileName += '.json';
    }

    // Final sanitization to ensure file name is valid
    return ValidationRules.sanitizeFileName(fileName);
  }

  /**
   * Formats a date as YYYY-MM-DD
   */
  private static formatDate(date: Date): string {
    const year = date.getFullYear();
    const month = ('0' + (date.getMonth() + 1)).slice(-2);
    const day = ('0' + date.getDate()).slice(-2);
    return `${year}-${month}-${day}`;
  }

  /**
   * Formats a time as HH-MM-SS
   */
  private static formatTime(date: Date): string {
    const hours = ('0' + date.getHours()).slice(-2);
    const minutes = ('0' + date.getMinutes()).slice(-2);
    const seconds = ('0' + date.getSeconds()).slice(-2);
    return `${hours}-${minutes}-${seconds}`;
  }

  /**
   * Generates a simple GUID
   */
  private static generateGuid(): string {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
      const r = (Math.random() * 16) | 0;
      const v = c === 'x' ? r : (r & 0x3) | 0x8;
      return v.toString(16);
    });
  }
}
