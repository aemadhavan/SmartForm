import * as React from 'react';
import { useEffect } from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { IMessageDisplayProps } from './IMessageDisplayProps';
import styles from './MessageDisplay.module.scss';

const MessageDisplay: React.FC<IMessageDisplayProps> = ({ message, messageType, onDismiss }) => {
  const getMessageBarType = (): MessageBarType => {
    switch (messageType) {
      case 'success':
        return MessageBarType.success;
      case 'error':
        return MessageBarType.error;
      case 'warning':
        return MessageBarType.warning;
      case 'info':
      default:
        return MessageBarType.info;
    }
  };

  // Auto-dismiss success messages after 5 seconds
  useEffect(() => {
    if (messageType === 'success') {
      const timer = setTimeout(() => {
        onDismiss();
      }, 5000);

      return () => clearTimeout(timer);
    }
  }, [messageType, onDismiss]);

  return (
    <div className={styles.messageDisplay}>
      <MessageBar
        messageBarType={getMessageBarType()}
        onDismiss={onDismiss}
        dismissButtonAriaLabel="Close"
        isMultiline={message.length > 100}
        role="alert"
        aria-live="polite"
      >
        {message}
      </MessageBar>
    </div>
  );
};

export default MessageDisplay;
