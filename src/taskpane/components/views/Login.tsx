import React, { useEffect, useId, useState } from 'react';
import {
  Input,
  Button,
  Title1,
  Body1,
  Caption1,
  Card,
  makeStyles,
  shorthands,
  tokens,
  Spinner,
  Toaster,
  useToastController,
  Toast,
  ToastTitle,
  ToastBody,
  ToastIntent,
  Link,
  FluentProvider,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogTitle,
  DialogBody,
  DialogActions,
  Field,
} from '@fluentui/react-components';
import { Settings24Regular } from '@fluentui/react-icons';

import auth from '../services/auth';
import { useNavigate } from 'react-router-dom';
import { mainSequenceTheme } from '../../theme';
import {
  clearApiBaseUrl,
  DEFAULT_API_BASE_URL,
  getApiBaseUrl,
  getApiBaseUrlSync,
  setApiBaseUrl,
} from '../../../shared/apiConfig';

const MainSequenceLogo = () => (
  <img
    src="https://main-sequence.app/static/media/logos/MS_logo_long_white.png"
    alt="MainSequence logo"
    style={{ maxWidth: "220px", height: "auto" }}
  />
);

const useStyles = makeStyles({
  container: {
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    minHeight: '100vh',
    background: `linear-gradient(135deg, ${tokens.colorNeutralBackground2}, ${tokens.colorNeutralBackground4})`,
  },
  card: {
    maxWidth: '400px',
    marginL: '10px',
    width: '100%',
    ...shorthands.borderRadius(tokens.borderRadiusXLarge),
    ...shorthands.padding('40px'),
    boxShadow: tokens.shadow16,
    backgroundColor: tokens.colorNeutralBackground1,

    // Media query for mobile screens
    '@media (max-width: 480px)': {
      ...shorthands.padding('24px'),
      ...shorthands.margin('10px'),
      // Reduce padding on smaller screens
    },
  },
  logoContainer: {
    display: 'flex',
    justifyContent: 'center',
    marginBottom: tokens.spacingVerticalL,
  },
  headerRow: {
    position: 'relative',
    marginBottom: tokens.spacingVerticalS,
  },
  headerText: {
    textAlign: 'center',
  },
  form: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalL),
  },
  title: {
    textAlign: 'center',
    justifyContent: 'center',
    width: '100%'
  },
  subtitle: {
    textAlign: 'center',
    color: tokens.colorNeutralForeground2,
    justifyContent: 'center',
    width: '100%'
  },
  button: {
    marginTop: tokens.spacingVerticalS,
  },
  forgotPassword: {
    textAlign: 'right',
    marginTop: `-${tokens.spacingVerticalM}`,
    marginBottom: tokens.spacingVerticalS,
  },
  footer: {
    marginTop: tokens.spacingVerticalXXL,
    textAlign: 'center',
    color: tokens.colorNeutralForeground2,
  },
  settingsButton: {
    position: 'absolute',
    top: 0,
    right: 0,
  },
  settingsContent: {
    display: 'flex',
    flexDirection: 'column',
    ...shorthands.gap(tokens.spacingVerticalM),
  },
  settingsHint: {
    color: tokens.colorNeutralForeground2,
  },
});

const Login = () => {
  const styles = useStyles();
  const toasterId = useId(); // Generate a unique ID for the toaster
  const { dispatchToast } = useToastController(toasterId);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [endpointInput, setEndpointInput] = useState(getApiBaseUrlSync());
  const [savingEndpoint, setSavingEndpoint] = useState(false);
  const navigate = useNavigate()
  // --- Best Practice: Reusable Notification Function ---
  const notify = ({ title, body, intent }: { title: string; body: string; intent: ToastIntent }) => {
    dispatchToast(
      <Toast>
        <ToastTitle>{title}</ToastTitle>
        <ToastBody>{body}</ToastBody>
      </Toast>,
      { intent }
    );
  };

  useEffect(() => {
    let isMounted = true;
    (async () => {
      const current = await getApiBaseUrl();
      if (isMounted) {
        setEndpointInput(current);
      }
    })();
    return () => {
      isMounted = false;
    };
  }, []);

  const handleSaveEndpoint = async () => {
    const trimmed = endpointInput.trim();
    if (!trimmed) {
      notify({
        title: 'Invalid Endpoint',
        body: 'Please enter a valid API endpoint.',
        intent: 'warning',
      });
      return;
    }

    setSavingEndpoint(true);
    try {
      const saved = await setApiBaseUrl(trimmed);
      setEndpointInput(saved);
      notify({
        title: 'Endpoint Saved',
        body: `Using ${saved}`,
        intent: 'success',
      });
      setSettingsOpen(false);
    } catch (err: unknown) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to save endpoint.';
      notify({
        title: 'Save Failed',
        body: errorMessage,
        intent: 'error',
      });
    } finally {
      setSavingEndpoint(false);
    }
  };

  const handleResetEndpoint = async () => {
    setSavingEndpoint(true);
    try {
      await clearApiBaseUrl();
      setEndpointInput(DEFAULT_API_BASE_URL);
      notify({
        title: 'Endpoint Reset',
        body: `Reverted to ${DEFAULT_API_BASE_URL}`,
        intent: 'info',
      });
    } catch (err: unknown) {
      const errorMessage = err instanceof Error ? err.message : 'Failed to reset endpoint.';
      notify({
        title: 'Reset Failed',
        body: errorMessage,
        intent: 'error',
      });
    } finally {
      setSavingEndpoint(false);
    }
  };

  const handleLogin = async (event: React.FormEvent) => {
    event.preventDefault();

    if (!email || !password) {
      notify({
        title: 'Missing Information',
        body: 'Please enter both your email and password.',
        intent: 'warning',
      });
      return;
    }

    setLoading(true);
    try {
      await new Promise(resolve => setTimeout(resolve, 1500));
      await auth(email, password, (result, error) => {
        if (result) {
          console.log('result', result);

          localStorage.setItem('token', result.access);
          localStorage.setItem('refresh_token', result.refresh);
          OfficeRuntime.storage.setItem('token', result.access);

          navigate('/Home')
          setLoading(false);
          notify({
            title: 'Login Successful',
            body: 'Welcome back! Redirecting you now...',
            intent: 'success',
          });
          localStorage.setItem('email', email)
        }
        if (error) {
          setLoading(false);
          notify({
            title: 'Login Failed',
            body: error.message || 'An error occurred during login. Please try again.',
            intent: 'error',
          });
        }
      });
    } catch (err: unknown) {
      const errorMessage = err instanceof Error ? err.message : 'An unknown error occurred.';
      notify({
        title: 'Login Failed',
        body: errorMessage,
        intent: 'error',
      });
    } finally {
      setLoading(false);
    }
  };

  return (
    <FluentProvider theme={mainSequenceTheme}>
      <div className={styles.container}>
        {/* Use the generated ID for the toaster instance */}
        <Toaster toasterId={toasterId} />
        <Card className={styles.card}>
          <div className={styles.logoContainer}>
            <MainSequenceLogo />
          </div>

          <div className={styles.headerRow}>
            <div className={styles.headerText}>
              <Title1 className={styles.title}>Sign In</Title1>
              <Body1 className={styles.subtitle}>Sign in with your Main Sequence credentials</Body1>
            </div>
            <Dialog open={settingsOpen} onOpenChange={(_, data) => setSettingsOpen(data.open)}>
              <DialogTrigger disableButtonEnhancement>
                <Button
                  appearance="subtle"
                  className={styles.settingsButton}
                  icon={<Settings24Regular />}
                  aria-label="Settings"
                  disabled={loading}
                />
              </DialogTrigger>
              <DialogSurface>
                <DialogBody>
                  <DialogTitle>Settings</DialogTitle>
                  <div className={styles.settingsContent}>
                    <Field label="API endpoint">
                      <Input
                        value={endpointInput}
                        onChange={(e) => setEndpointInput(e.target.value)}
                        placeholder={DEFAULT_API_BASE_URL}
                        disabled={savingEndpoint}
                      />
                    </Field>
                    <Caption1 className={styles.settingsHint}>
                      Default: {DEFAULT_API_BASE_URL}
                    </Caption1>
                  </div>
                </DialogBody>
                <DialogActions>
                  <Button appearance="secondary" onClick={handleResetEndpoint} disabled={savingEndpoint}>
                    Reset
                  </Button>
                  <Button appearance="primary" onClick={handleSaveEndpoint} disabled={savingEndpoint}>
                    Save
                  </Button>
                </DialogActions>
              </DialogSurface>
            </Dialog>
          </div>

          <form onSubmit={handleLogin} className={styles.form}>
            <Input
              required
              type="email"
              placeholder="Enter your email"
              value={email}
              size="large"
              onChange={(e) => setEmail(e.target.value)}
              // --- UX Improvement: Disable input during loading ---
              disabled={loading}
            />
            <div>
              <Input
                required
                type="password"
                placeholder="Enter your password"
                value={password}
                size="large"
                onChange={(e) => setPassword(e.target.value)}
                style={{ width: '100%' }}
                disabled={loading}
              />
              <div className={styles.forgotPassword}>
                <Link href="#" disabled={loading}>
                  Forgot Password?
                </Link>
              </div>
            </div>

            <Button
              appearance="primary"
              className={styles.button}
              type="submit"
              disabled={loading}
              size="large"
              icon={loading ? <Spinner size="tiny" /> : null}
            >
              {loading ? 'Signing In...' : 'Sign In'}
            </Button>
            {/* <Button
            appearance="primary"
            className={styles.button}
            type="submit"
            // disabled={loading}
            size="large"
            icon={loading ? <Spinner size="tiny" /> : null}
            onClick={() => {
              refresh(OfficeRuntime.storage.getItem('refresh'), (result, err) => {
                if (result) {
                  localStorage.setItem('token', JSON.parse(result).access);
                  localStorage.setItem('refresh', JSON.parse(result).refresh);
                  OfficeRuntime.storage.setItem('token', JSON.parse(result).access);
                  OfficeRuntime.storage.setItem('refresh', JSON.parse(result).refresh);

                  console.log('refreshed', result);
                }
                if (err) {

                  console.log('failed refresh', err);
                }
              })
            }}
          >
            refresh
          </Button>

          <Button
            appearance="primary"
            className={styles.button}
            type="submit"
            // disabled={loading}
            size="large"
            icon={loading ? <Spinner size="tiny" /> : null}
            onClick={() => {
              FetchData(
                "1/1/2000",
                null,
                [],
                true,
                true,
                (result, error) => {
                  if (result) {
                    console.log("Results:", result.results);
                  } else {
                    console.error("Error:", error);
                  }
                }
              );
            }}
          >
            fetch data
          </Button> */}
          </form>

          <div className={styles.footer}>
            {/* <Caption1>
            Don't have an account?{' '}
            <Link href="#" disabled={loading}>
              Sign Up
            </Link>
          </Caption1> */}
          </div>
        </Card>
      </div>
    </FluentProvider>
  );
};

export default Login;
