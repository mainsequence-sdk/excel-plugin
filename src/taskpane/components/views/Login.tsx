import React, { useState, useId } from 'react';
import {
  Input,
  Button,
  Title1,
  Body1,
  Caption1,
  Card,
  CardHeader,
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
} from '@fluentui/react-components';

import auth from '../services/auth';
import { useNavigate } from 'react-router-dom';
import { mainSequenceTheme } from '../../theme';

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
});

const Login = () => {
  const styles = useStyles();
  const toasterId = useId(); // Generate a unique ID for the toaster
  const { dispatchToast } = useToastController(toasterId);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loading, setLoading] = useState(false);
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
          localStorage.setItem('refresh', result.refresh);
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
            body: error,
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

          <CardHeader
            header={<Title1 className={styles.title}>Sign In</Title1>}
            description={<Body1 className={styles.subtitle}>Sign in with your Main Sequence credentials</Body1>}
          />

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
