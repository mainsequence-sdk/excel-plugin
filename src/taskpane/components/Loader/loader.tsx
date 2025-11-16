import * as React from "react";
import { Spinner, makeStyles } from "@fluentui/react-components";

const useStyles = makeStyles({
    overlay: {
        position: "fixed",
        top: 0,
        left: 0,
        width: "100vw",
        height: "100vh",
        backgroundColor: "rgba(255, 255, 255, 0.8)",
        display: "flex",
        justifyContent: "center",
        alignItems: "center",
        zIndex: 9999,
    },
    greenProgress: {
        "--colorBrandStroke1": "#28a745",
        "--colorBrandStroke2": "#c6f2c8",
        "--colorBrandStroke2Contrast": "#c6f2c8"
    },
});

const Loader: React.FC = () => {
    const styles = useStyles();

    return (
        <div className={styles.overlay}>
            <Spinner size="medium" className={styles.greenProgress} />
        </div>
    );
};

export default Loader;