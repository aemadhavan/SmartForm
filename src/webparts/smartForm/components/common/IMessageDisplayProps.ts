export interface IMessageDisplayProps {
  message: string;
  messageType: 'success' | 'error' | 'info' | 'warning';
  onDismiss: () => void;
}
