import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Dialog, DialogType } from '@fluentui/react/lib/Dialog';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import { Icon } from '@fluentui/react/lib/Icon';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { useTheme, mergeStyles } from '@fluentui/react';
import { Log } from '@microsoft/sp-core-library';
import * as strings from 'PdfExportCommandSetStrings';
import styles from './ProgressDialog.module.scss';

interface IProgressDialogState {
  isVisible: boolean;
  title: string;
  message: string;
  current: number;
  total: number;
  currentFileName: string;
  status: 'processing' | 'completed' | 'error';
}

type ProgressDialogAction =
  | { type: 'SHOW_DIALOG'; title: string; message: string; total: number }
  | { type: 'UPDATE_PROGRESS'; current: number; total: number; fileName: string; status: 'processing' | 'completed' | 'error' }
  | { type: 'CLOSE_DIALOG' };

const progressDialogReducer = (state: IProgressDialogState, action: ProgressDialogAction): IProgressDialogState => {
  switch (action.type) {
    case 'SHOW_DIALOG':
      return {
        ...state,
        isVisible: true,
        title: action.title,
        message: action.message,
        total: action.total,
        current: 0,
        currentFileName: '',
        status: 'processing'
      };
    case 'UPDATE_PROGRESS':
      return {
        ...state,
        current: action.current,
        total: action.total,
        currentFileName: action.fileName,
        status: action.status
      };
    case 'CLOSE_DIALOG':
      return {
        ...state,
        isVisible: false,
        current: 0,
        total: 0,
        currentFileName: '',
        status: 'processing'
      };
    default:
      return state;
  }
};

interface IProgressDialogContentProps {
  title: string;
  message: string;
  current: number;
  total: number;
  currentFileName: string;
  status: 'processing' | 'completed' | 'error';
  hidden: boolean;
}

const StatusIcon: React.FC<{ status: 'processing' | 'completed' | 'error' }> = ({ status }) => {
  const theme = useTheme();
  
  const iconStyle = {
    fontSize: '16px',
    marginRight: '8px'
  };
  
  switch (status) {
    case 'completed':
      return <Icon iconName="CheckMark" style={{ ...iconStyle, color: theme.palette.green }} />;
    case 'error':
      return <Icon iconName="ErrorBadge" style={{ ...iconStyle, color: theme.palette.redDark }} />;
    default:
      return <Icon iconName="Processing" style={{ ...iconStyle, color: theme.palette.themePrimary }} />;
  }
};

const ProgressDialogContent: React.FC<IProgressDialogContentProps> = ({
  title,
  message,
  current,
  total,
  currentFileName,
  status,
  hidden
}) => {
  const theme = useTheme();
  const progressPercentage = total > 0 ? (current / total) : 0;
  const progressText = `${current} ${strings.Of} ${total} ${strings.FilesProcessed}`;

  const dialogStyles = mergeStyles({
    selectors: {
      '.ms-Dialog-main': {
        backgroundColor: theme.palette.white,
        borderRadius: '8px',
        boxShadow: theme.effects.elevation16,
        border: `1px solid ${theme.palette.neutralLight}`,
        minWidth: '480px',
        maxWidth: '600px',
        overflow: 'hidden'
      },
      '.ms-Dialog-title': {
        fontSize: theme.fonts.large.fontSize,
        fontWeight: '600',
        color: theme.palette.neutralPrimary,
        padding: '24px 24px 16px 24px',
        margin: '0',
        borderBottom: `1px solid ${theme.palette.neutralLighter}`
      },
      '.ms-Dialog-content': {
        padding: '0'
      }
    }
  });

  return (
    <div className={`${styles['progress-dialog-container']} ${dialogStyles}`}>
      <Dialog
        hidden={hidden}
        dialogContentProps={{ 
          type: DialogType.normal, 
          title: title,
          showCloseButton: false,
          className: undefined
        }}
        modalProps={{ 
          isDarkOverlay: true, 
          isBlocking: true,
          dragOptions: undefined,
          className: undefined
        }}
        minWidth={480}
        maxWidth={600}
      >
        <div className={styles['progress-body']} style={{ padding: '32px' }}>
          {/* Progress Header */}
          <Stack tokens={{ childrenGap: 20 }} horizontalAlign="center">
            <div className={styles['progress-icon-container']} style={{ backgroundColor: theme.palette.themeLighter }}>
              <Icon iconName="CloudDownload" className={styles['progress-header-icon']} style={{ color: theme.palette.themePrimary, fontSize: 32 }} />
            </div>
            <Stack tokens={{ childrenGap: 8 }} horizontalAlign="center">
              <Text variant="large" className={styles['progress-header-title']} style={{ color: theme.palette.neutralPrimary, fontWeight: 600 }}>
                {message}
              </Text>
              <Text variant="medium" className={styles['progress-header-subtitle']} style={{ color: theme.palette.neutralSecondary }}>
                Converting documents to PDF
              </Text>
            </Stack>
          </Stack>

          {/* Main Progress Section */}
          <Stack tokens={{ childrenGap: 20 }} style={{ marginTop: '24px' }}>
            <ProgressIndicator
              label={progressText}
              description={`${Math.round(progressPercentage * 100)}%`}
              percentComplete={progressPercentage}
              className={styles['main-progress-indicator']}
            />
            
            {/* Current File Status */}
            {currentFileName && (
              <Stack tokens={{ childrenGap: 8 }}>
                <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                  {status === 'processing' && strings.Processing}
                  {status === 'completed' && strings.Completed}
                  {status === 'error' && strings.Error}
                </Text>
                <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center" className={styles['file-card']}>
                  <StatusIcon status={status} />
                  <Text variant="medium" style={{ color: theme.palette.neutralPrimary, flex: 1 }}>
                    {currentFileName}
                  </Text>
                </Stack>
              </Stack>
            )}

            {/* Status Message */}
            {status === 'error' && currentFileName && (
              <MessageBar
                messageBarType={MessageBarType.warning}
                className={styles['status-message']}
              >
                <Text variant="small">
                  Some files could not be converted. The process will continue with remaining files.
                </Text>
              </MessageBar>
            )}
          </Stack>
        </div>
      </Dialog>
    </div>
  );
};

class ProgressDialogController {
  private container: HTMLElement;
  private dispatch!: React.Dispatch<ProgressDialogAction>;

  constructor() {
    this.container = document.createElement('div');
    document.body.appendChild(this.container);

    const Wrapper: React.FC = () => {
      const [state, dispatch] = React.useReducer(progressDialogReducer, {
        isVisible: false,
        title: '',
        message: '',
        current: 0,
        total: 0,
        currentFileName: '',
        status: 'processing'
      });

      this.dispatch = dispatch;

      return (
        <ProgressDialogContent
          title={state.title}
          message={state.message}
          current={state.current}
          total={state.total}
          currentFileName={state.currentFileName}
          status={state.status}
          hidden={!state.isVisible}
        />
      );
    };

    ReactDOM.render(<Wrapper />, this.container);
  }

  public show(title: string, message: string, total: number): void {
    this.dispatch({ type: 'SHOW_DIALOG', title, message, total });
    Log.info('ProgressDialogController', `Showing progress dialog: ${title} - ${message} (Total: ${total})`);
  }

  public updateProgress(current: number, total: number, fileName: string, status: 'processing' | 'completed' | 'error'): void {
    this.dispatch({ type: 'UPDATE_PROGRESS', current, total, fileName, status });
    Log.info('ProgressDialogController', `Progress update: ${current}/${total} - ${fileName} - ${status}`);
  }

  public close(): void {
    this.dispatch({ type: 'CLOSE_DIALOG' });
    Log.info('ProgressDialogController', 'Closing progress dialog');
  }
}

export default new ProgressDialogController();