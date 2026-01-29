import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { Dialog, DialogType } from '@fluentui/react/lib/Dialog';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { ProgressIndicator } from '@fluentui/react/lib/ProgressIndicator';
import { Log } from '@microsoft/sp-core-library';
import { useTheme, mergeStyles } from '@fluentui/react';
import styles from './WaitDialog.module.scss';

interface IWaitDialogState {
  isVisible: boolean;
  title: string;
  message: string;
  error: string;
  showClose: boolean;
  showConfirm: boolean;
  confirmResolve?: (value: boolean) => void;
  progress?: number;
  currentFile?: string;
  estimatedTime?: string;
}

type WaitDialogAction =
  | { type: 'SHOW_DIALOG'; title: string; message: string }
  | { type: 'SHOW_ERROR'; title: string; message: string }
  | { type: 'SHOW_CONFIRM'; title: string; message: string; resolve: (value: boolean) => void }
  | { type: 'UPDATE_PROGRESS'; progress?: number; currentFile?: string; estimatedTime?: string }
  | { type: 'CLOSE_DIALOG' }
  | { type: 'CONFIRM_RESPONSE'; confirmed: boolean };

const waitDialogReducer = (state: IWaitDialogState, action: WaitDialogAction): IWaitDialogState => {
  switch (action.type) {
    case 'SHOW_DIALOG':
      return { 
        ...state, 
        isVisible: true, 
        title: action.title, 
        message: action.message, 
        error: '', 
        showClose: false,
        showConfirm: false,
        progress: undefined,
        currentFile: undefined,
        estimatedTime: undefined
      };
    case 'SHOW_ERROR':
      return { 
        ...state, 
        isVisible: true, 
        title: action.title, 
        message: action.message, 
        error: action.message, 
        showClose: true,
        showConfirm: false,
        progress: undefined,
        currentFile: undefined,
        estimatedTime: undefined
      };
    case 'SHOW_CONFIRM':
      return {
        ...state,
        isVisible: true,
        title: action.title,
        message: action.message,
        error: '',
        showClose: false,
        showConfirm: true,
        confirmResolve: action.resolve,
        progress: undefined,
        currentFile: undefined,
        estimatedTime: undefined
      };
    case 'CONFIRM_RESPONSE':
      if (state.confirmResolve) {
        state.confirmResolve(action.confirmed);
      }
      return {
        ...state,
        isVisible: false,
        showConfirm: false,
        confirmResolve: undefined,
        message: '',
        title: '',
        error: ''
      };
    case 'UPDATE_PROGRESS':
      return {
        ...state,
        progress: action.progress,
        currentFile: action.currentFile,
        estimatedTime: action.estimatedTime
      };
    case 'CLOSE_DIALOG':
      return { 
        ...state, 
        isVisible: false, 
        message: '', 
        title: '', 
        error: '', 
        showClose: false,
        showConfirm: false,
        confirmResolve: undefined,
        progress: undefined,
        currentFile: undefined,
        estimatedTime: undefined
      };
    default:
      return state;
  }
};

interface IWaitDialogContentProps {
  message: string;
  error: string;
  title: string;
  showClose: boolean;
  showConfirm: boolean;
  hidden: boolean;
  progress?: number;
  currentFile?: string;
  estimatedTime?: string;
  closeCallback: () => void;
  confirmCallback: (confirmed: boolean) => void;
}

const ErrorContent: React.FC<{ error: string }> = ({ error }) => {
  const theme = useTheme();
  
  return error ? (
    <Stack tokens={{ childrenGap: 20 }} horizontalAlign="center" className={styles['error-container']}>
      <div className={styles['error-icon-wrapper']}>
        <Icon iconName="ErrorBadge" className={styles['error-icon']} />
      </div>
      <Text variant="medium" className={styles['error-title']} style={{ color: theme.palette.neutralPrimary }}>
        {error}
      </Text>
    </Stack>
  ) : null;
};

const LoadingContent: React.FC<{ 
  message: string; 
  title: string; 
  progress?: number; 
  currentFile?: string; 
  estimatedTime?: string; 
}> = ({ message, title, progress, currentFile, estimatedTime }) => {
  const theme = useTheme();
  
  return (
    <Stack tokens={{ childrenGap: 24 }} className={styles['loading-container']}>
      {/* Header */}
      <Stack tokens={{ childrenGap: 12 }} horizontalAlign="center">
        <div className={styles['icon-wrapper']} style={{ backgroundColor: theme.palette.themeLighter }}>
          <Icon iconName="PDF" className={styles['main-icon']} style={{ color: theme.palette.themePrimary }} />
        </div>
        <Text variant="large" className={styles['main-title']} style={{ color: theme.palette.neutralPrimary }}>
          {title}
        </Text>
        <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
          <Spinner size={SpinnerSize.small} />
          <Text variant="medium" style={{ color: theme.palette.neutralSecondary }}>
            {message}
          </Text>
        </Stack>
        {estimatedTime && (
          <Text variant="small" style={{ color: theme.palette.themePrimary }}>
            {estimatedTime}
          </Text>
        )}
      </Stack>

      {/* Progress */}
      <div className={styles['progress-section']}>
        {progress !== undefined ? (
          <ProgressIndicator
            percentComplete={progress / 100}
            description={`${Math.round(progress)}%`}
            className={styles['main-progress']}
          />
        ) : (
          <ProgressIndicator
            description="Processing..."
            className={styles['main-progress']}
          />
        )}
      </div>

      {/* Current File */}
      {currentFile && (
        <Stack tokens={{ childrenGap: 8 }} className={styles['current-file-section']}>
          <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
            Processing:
          </Text>
          <Stack horizontal tokens={{ childrenGap: 12 }} verticalAlign="center" className={styles['file-card']}>
            <Icon iconName="Document" style={{ color: theme.palette.themePrimary, fontSize: 20 }} />
            <Stack tokens={{ childrenGap: 4 }} style={{ flex: 1 }}>
              <Text variant="medium" style={{ color: theme.palette.neutralPrimary }}>
                {currentFile}
              </Text>
              <Text variant="small" style={{ color: theme.palette.neutralSecondary }}>
                Converting to PDF...
              </Text>
            </Stack>
            <Spinner size={SpinnerSize.small} />
          </Stack>
        </Stack>
      )}
    </Stack>
  );
};

const ConfirmContent: React.FC<{ message: string; onConfirm: (confirmed: boolean) => void }> = ({ message, onConfirm }) => {
  const theme = useTheme();
  
  return (
    <Stack tokens={{ childrenGap: 24 }} horizontalAlign="center">
      <div className={styles['error-icon-wrapper']} style={{ backgroundColor: theme.palette.themeLighter }}>
        <Icon iconName="Warning" className={styles['error-icon']} style={{ color: theme.palette.yellow, fontSize: 32 }} />
      </div>
      <Text variant="medium" style={{ color: theme.palette.neutralPrimary, textAlign: 'center' }}>
        {message}
      </Text>
      <Stack horizontal tokens={{ childrenGap: 12 }} style={{ marginTop: '8px' }}>
        <DefaultButton
          text="Yes, Overwrite"
          onClick={() => onConfirm(true)}
          styles={{
            root: {
              backgroundColor: theme.palette.themePrimary,
              borderColor: theme.palette.themePrimary,
              color: theme.palette.white,
              minWidth: '120px',
              borderRadius: '2px'
            },
            rootHovered: {
              backgroundColor: theme.palette.themeDark,
              borderColor: theme.palette.themeDark
            }
          }}
        />
        <DefaultButton
          text="Cancel"
          onClick={() => onConfirm(false)}
          styles={{
            root: {
              backgroundColor: theme.palette.neutralLighter,
              borderColor: theme.palette.neutralLight,
              color: theme.palette.neutralPrimary,
              minWidth: '120px',
              borderRadius: '2px'
            },
            rootHovered: {
              backgroundColor: theme.palette.neutralLight,
              borderColor: theme.palette.neutralTertiary
            }
          }}
        />
      </Stack>
    </Stack>
  );
};

const WaitDialogContent: React.FC<IWaitDialogContentProps> = ({
  message,
  error,
  title,
  showClose,
  showConfirm,
  hidden,
  progress,
  currentFile,
  estimatedTime,
  closeCallback,
  confirmCallback,
}) => {
  const theme = useTheme();
  const dialogType = showClose ? DialogType.close : DialogType.normal;
  
  const dialogStyles = mergeStyles({
    selectors: {
      '.ms-Dialog-main': {
        backgroundColor: theme.palette.white,
        borderRadius: '8px',
        boxShadow: theme.effects.elevation16,
        border: `1px solid ${theme.palette.neutralLight}`,
        minWidth: '420px',
        maxWidth: '500px',
        overflow: 'hidden'
      },
      '.ms-Dialog-title': {
        display: 'none'
      },
      '.ms-Dialog-content': {
        padding: '0'
      }
    }
  });

  return (
    <div className={`${styles['dialog-container']} ${dialogStyles}`}>
      <Dialog
        hidden={hidden}
        dialogContentProps={{ 
          type: dialogType,
          showCloseButton: showClose,
          className: undefined
        }}
        modalProps={{ 
          isDarkOverlay: true, 
          isBlocking: !showClose,
          onDismiss: showClose ? closeCallback : undefined,
          className: undefined
        }}
        minWidth={420}
        maxWidth={500}
      >
        <div className={styles['dialog-body']} style={{ padding: '32px' }}>
          {showConfirm ? (
            <ConfirmContent message={message} onConfirm={confirmCallback} />
          ) : error ? (
            <ErrorContent error={error} />
          ) : (
            <LoadingContent 
              message={message} 
              title={title} 
              progress={progress}
              currentFile={currentFile}
              estimatedTime={estimatedTime}
            />
          )}
          
          {showClose && (
            <Stack horizontalAlign="center" style={{ marginTop: '24px' }}>
              <DefaultButton 
                text="Close" 
                onClick={closeCallback}
                styles={{
                  root: {
                    backgroundColor: theme.palette.themePrimary,
                    borderColor: theme.palette.themePrimary,
                    color: theme.palette.white,
                    minWidth: '100px',
                    borderRadius: '2px'
                  },
                  rootHovered: {
                    backgroundColor: theme.palette.themeDark,
                    borderColor: theme.palette.themeDark
                  }
                }}
              />
            </Stack>
          )}
        </div>
      </Dialog>
    </div>
  );
};

const dialogDiv = document.createElement('div');
document.body.appendChild(dialogDiv);

const WaitDialog: React.FC<IWaitDialogState & { onClose: () => void; onConfirm: (confirmed: boolean) => void }> = ({
  isVisible,
  message,
  title,
  error,
  showClose,
  showConfirm,
  progress,
  currentFile,
  estimatedTime,
  onClose,
  onConfirm,
}) => {
  return ReactDOM.createPortal(
    <WaitDialogContent
      message={message}
      title={title}
      error={error}
      showClose={showClose}
      showConfirm={showConfirm}
      progress={progress}
      currentFile={currentFile}
      estimatedTime={estimatedTime}
      hidden={!isVisible}
      closeCallback={onClose}
      confirmCallback={onConfirm}
    />,
    dialogDiv
  );
};

class WaitDialogController {
  private container: HTMLElement;
  private dispatch!: React.Dispatch<WaitDialogAction>;

  constructor() {
    this.container = document.createElement('div');
    document.body.appendChild(this.container);

    const Wrapper: React.FC = () => {
      const [state, dispatch] = React.useReducer(waitDialogReducer, {
        isVisible: false,
        message: '',
        title: '',
        error: '',
        showClose: false,
        showConfirm: false,
      });

      this.dispatch = dispatch;

      return (
        <WaitDialog 
          {...state} 
          onClose={() => dispatch({ type: 'CLOSE_DIALOG' })}
          onConfirm={(confirmed) => dispatch({ type: 'CONFIRM_RESPONSE', confirmed })}
        />
      );
    };

    ReactDOM.render(<Wrapper />, this.container);
  }

  public show(title: string, message: string) {
    this.dispatch({ type: 'SHOW_DIALOG', title, message });
    Log.info('WaitDialogController', `Showing dialog: ${title} - ${message}`);
  }

  public showError(title: string, message: string) {
    this.dispatch({ type: 'SHOW_ERROR', title, message });
    Log.error('WaitDialogController', new Error(`Showing error dialog: ${title} - ${message}`));
  }

  public updateProgress(progress?: number, currentFile?: string, estimatedTime?: string) {
    this.dispatch({ type: 'UPDATE_PROGRESS', progress, currentFile, estimatedTime });
    Log.info('WaitDialogController', `Updating progress: ${progress}% - ${currentFile}`);
  }

  public close() {
    this.dispatch({ type: 'CLOSE_DIALOG' });
    Log.info('WaitDialogController', 'Closing dialog.');
  }

  public async confirm(title: string, message: string): Promise<boolean> {
    return new Promise<boolean>((resolve) => {
      this.dispatch({ type: 'SHOW_CONFIRM', title, message, resolve });
      Log.info('WaitDialogController', `Showing confirmation: ${title} - ${message}`);
    });
  }
}

export default new WaitDialogController();