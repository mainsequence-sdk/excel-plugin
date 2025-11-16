// import React, { useState, useEffect } from "react";
// import {
//   FluentProvider,
//   teamsLightTheme,
//   Button,
//   Text,
//   Title3,
//   Subtitle1,
//   Dialog,
//   DialogTrigger,
//   DialogSurface,
//   DialogBody,
//   DialogTitle,
//   DialogContent,
//   DialogActions,
//   Avatar,
//   Divider,
//   makeStyles,
//   tokens,
//   Spinner,
// } from "@fluentui/react-components";
// import { SignOut20Regular } from "@fluentui/react-icons";
// import { useNavigate } from "react-router-dom";
// const MainSequenceLogo = () => (
//   <svg width="64" height="64" viewBox="0 0 24 24" fill="none" xmlns="https://dev-tsorm.ngrok.app/static/media/logos/emblem.png">
//     <path
//       d="M12 2L2 7L12 12L22 7L12 2Z"
//       stroke={tokens.colorBrandForeground1}
//       strokeWidth="2"
//       strokeLinecap="round"
//       strokeLinejoin="round"
//     />
//     <path
//       d="M2 17L12 22L22 17"
//       stroke={tokens.colorBrandForeground1}
//       strokeWidth="2"
//       strokeLinecap="round"
//       strokeLinejoin="round"
//     />
//     <path
//       d="M2 12L12 17L22 12"
//       stroke={tokens.colorBrandForeground1}
//       strokeWidth="2"
//       strokeLinecap="round"
//       strokeLinejoin="round"
//     />
//   </svg>
// );
// const useStyles = makeStyles({
//   container: {
//     width: "100%",
//     maxWidth: "360px",
//     margin: "0 auto",
//     padding: "0",
//     backgroundColor: tokens.colorNeutralBackground1,
//     minHeight: "98vh",
//     display: "flex",
//     flexDirection: "column",
//   },
//   navbar: {
//     display: "flex",
//     alignItems: "center",
//     justifyContent: "space-between",
//     padding: "12px 16px",
//     borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
//   },
//   brand: {
//     fontWeight: 600,
//     fontSize: "16px",
//     color: tokens.colorBrandForegroundLink,
//   },
//   body: {
//     padding: "20px",
//     display: "flex",
//     flexDirection: "column",
//     gap: "16px",
//   },
//   section: {
//     display: "flex",
//     flexDirection: "column",
//     gap: "8px",
//     textAlign: "center",
//   },
//   footer: {
//     marginTop: "auto",
//     textAlign: "center",
//     padding: "8px",
//     fontSize: "12px",
//     color: tokens.colorNeutralForeground3,
//   },
// });

// export default function MainSequenceDashboard() {
//   const styles = useStyles();
//   const [loading, setLoading] = useState(true);
//   const [user, setUser] = useState({
//     name: localStorage.getItem("email"),
//     signedIn: false,
//   });
// const navigate=useNavigate()
//   useEffect(() => {
//     const checkToken = async () => {
//       const token = await OfficeRuntime.storage.getItem("token");
//       setUser((u) => ({ ...u, signedIn: !!token }));
//       setLoading(false);
//     };
//     checkToken();
//   }, []);

//   const handleLogout = async () => {
//     await OfficeRuntime.storage.removeItem("token");
//     setUser((u) => ({ ...u, signedIn: false }));
//     alert("Signed out successfully.");
//   };

//   if (loading) {
//     return (
//       <FluentProvider theme={teamsLightTheme}>
//         <div className={styles.container}>
//           <div className={styles.body}>
//             <Spinner label="Checking sign-in status..." />
//           </div>
//         </div>
//       </FluentProvider>
//     );
//   }

//   return (
//     <FluentProvider theme={teamsLightTheme}>
//       <div className={styles.container}>
//         {/* Navbar */}
//         <div className={styles.navbar}>

//           <Text className={styles.brand}>MainSequence</Text>

//           <Dialog>
//             <DialogTrigger disableButtonEnhancement>
//               <Avatar
//                 name={user.name || "User"}
//                 color="brand"
//                 aria-label="Profile"
//                 style={{ cursor: "pointer" }}
//               />
//             </DialogTrigger>

//             <DialogSurface>
//               <DialogBody style={{ margin: "8px" }}>
//                 <DialogTitle>Profile</DialogTitle>
//                 <DialogContent>
//                   <Text>
//                     {user.signedIn
//                       ? "✅ You are currently signed in."
//                       : "⚠️ Your session has expired."}
//                   </Text>
//                   <Divider style={{ margin: "8px 0" }} />
//                   <Text weight="semibold">{user.name || "Unknown user"}</Text>
//                 </DialogContent>
//                 <DialogActions>
//                   {user.signedIn && (
//                     <Button
//                       icon={<SignOut20Regular />}
//                       appearance="primary"
//                       onClick={handleLogout}
//                     >
//                       Log out
//                     </Button>
//                   )}
//                   <DialogTrigger disableButtonEnhancement>
//                     <Button appearance="secondary">Close</Button>
//                   </DialogTrigger>
//                 </DialogActions>
//               </DialogBody>
//             </DialogSurface>
//           </Dialog>
//         </div>


//         {/* Body */}
//         <div className={styles.body}>
//           {user.signedIn ? (
//             <div className={styles.section}>
//               <MainSequenceLogo />
//               <Title3>Welcome to MainSequence</Title3>
//               <Subtitle1>You've successfully signed in.</Subtitle1>
//               <Text>
//                 You can now use the add-in features to fetch and work with your
//                 data inside Excel.
//               </Text>
//             </div>
//           ) : (
//             <div className={styles.section}>
//               <Title3>Session Expired</Title3>
//               <Subtitle1>Please sign in again to continue.</Subtitle1>
//               <Button appearance="primary" onClick={() =>navigate('/')}>
//                 Sign In Again
//               </Button>
//             </div>
//           )}
//         </div>

//         <div className={styles.footer}>
//           <Text>© 2025 MainSequence Add-in</Text>
//         </div>
//       </div>
//     </FluentProvider>
//   );
// }










import React, { useState, useEffect } from "react";
import {
  FluentProvider,
  Button,
  Text,
  Title3,
  Subtitle1,
  Dialog,
  DialogTrigger,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Avatar,
  Divider,
  makeStyles,
  tokens,
  Spinner,
  Card,
  CardHeader,
  CardPreview,
  CardFooter,
} from "@fluentui/react-components";
import { SignOut20Regular } from "@fluentui/react-icons";
import { useNavigate } from "react-router-dom";
import { mainSequenceTheme } from "../../theme";

const MainSequenceLogo = () => (
  <img
    src="https://main-sequence.app/static/media/logos/MS_logo_long_white.png"
    alt="MainSequence logo"
    style={{ maxWidth: "200px", height: "auto" }}
  />
);

const useStyles = makeStyles({
  container: {
    width: "100%",
    height: "100vh",
    backgroundColor: tokens.colorNeutralBackground2,
    display: "flex",
    flexDirection: "column",
  },
  navbar: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    padding: "12px 20px",
    borderBottom: `1px solid ${tokens.colorNeutralStroke1}`,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow4,
  },
  brand: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    fontWeight: 600,
    fontSize: "18px",
    color: tokens.colorBrandForegroundLink,
  },
  body: {
    flexGrow: 1,
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    padding: "32px",
  },
  card: {
    width: "100%",
    maxWidth: "400px",
    padding: "32px",
    borderRadius: tokens.borderRadiusXLarge,
    backgroundColor: tokens.colorNeutralBackground1,
    boxShadow: tokens.shadow8,
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    gap: "16px",
    textAlign: "center",
  },
  footer: {
    textAlign: "center",
    padding: "12px 0",
    fontSize: "12px",
    color: tokens.colorNeutralForeground3,
  },
  logoWrapper: {
    transition: "transform 0.3s ease",
    ":hover": {
      transform: "scale(1.05)",
    },
  },
  spinnerContainer: {
    display: "flex",
    justifyContent: "center",
    alignItems: "center",
    height: "100vh",
  },
  avatar: {
    backgroundColor: "#468DFF",
    color: tokens.colorNeutralForegroundOnBrand,
  },
});

export default function MainSequenceDashboard() {
  const styles = useStyles();
  const [loading, setLoading] = useState(true);
  const [user, setUser] = useState({
    name: localStorage.getItem("email"),
    signedIn: false,
  });
  const navigate = useNavigate();

  useEffect(() => {
    const checkToken = async () => {
      const token = await OfficeRuntime.storage.getItem("token");
      setUser((u) => ({ ...u, signedIn: !!token }));
      setLoading(false);
    };
    checkToken();
  }, []);

  const handleLogout = async () => {
    await OfficeRuntime.storage.removeItem("token");
    setUser((u) => ({ ...u, signedIn: false }));
    alert("Signed out successfully.");
  };

  if (loading) {
    return (
      <FluentProvider theme={mainSequenceTheme}>
        <div className={styles.spinnerContainer}>
          <Spinner label="Checking sign-in status..." />
        </div>
      </FluentProvider>
    );
  }

  return (
    <FluentProvider theme={mainSequenceTheme}>
      <div className={styles.container}>
        {/* Navbar */}
        <div className={styles.navbar}>
          <div className={styles.brand}>
            <MainSequenceLogo />
          </div>

          <Dialog>
            <DialogTrigger disableButtonEnhancement>
              <Avatar
                name={user.name || "User"}
                className={styles.avatar}
                aria-label="Profile"
                style={{ cursor: "pointer" }}
              />
            </DialogTrigger>

            <DialogSurface>
              <DialogBody>
                <DialogTitle>Profile</DialogTitle>
                <DialogContent>
                  <Text>
                    {user.signedIn
                      ? "✅ You are currently signed in."
                      : "⚠️ Your session has expired."}
                  </Text>
                  <Divider style={{ margin: "10px 0" }} />
                  <Text weight="semibold">{user.name || "Unknown user"}</Text>
                </DialogContent>
                <DialogActions>
                  {user.signedIn && (
                    <Button
                      icon={<SignOut20Regular />}
                      appearance="primary"
                      onClick={handleLogout}
                    >
                      Log out
                    </Button>
                  )}
                  <DialogTrigger disableButtonEnhancement>
                    <Button appearance="secondary">Close</Button>
                  </DialogTrigger>
                </DialogActions>
              </DialogBody>
            </DialogSurface>
          </Dialog>
        </div>

        {/* Main Content */}
        <div className={styles.body}>
          <div className={styles.card}>
            <div className={styles.logoWrapper}>
              <MainSequenceLogo />
            </div>
            {user.signedIn ? (
              <>
                <Title3 style={{textAlign:'center'}}>Welcome</Title3>
                <Subtitle1 style={{textAlign:'center'}}>You’ve successfully signed in.</Subtitle1>
                <Text style={{textAlign:'center'}}>
                  You can now access all available tools and data inside Excel.
                </Text>
              </>
            ) : (
              <>
                <Title3>Session Expired</Title3>
                <Subtitle1>Please sign in again to continue.</Subtitle1>
                <Button
                  appearance="primary"
                  onClick={() => navigate("/")}
                  style={{ marginTop: "8px" }}
                >
                  Sign In Again
                </Button>
              </>
            )}
          </div>
        </div>

        {/* Footer */}
        <div className={styles.footer}>
          <Text>© 2025 MainSequence Add-in</Text>
        </div>
      </div>
    </FluentProvider>
  );
}
