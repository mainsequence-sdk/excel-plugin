

// export default RouterApp;
import React from "react";
import { HashRouter as Router, Routes, Route } from "react-router-dom";
import Login from "../views/Login";
import Home from "../views/Home";

const RouterApp: React.FC = () => {

    return (
        <>
            <Router>
                <Routes>
                    
                    <Route path="/" element={<Login />} />
                    <Route path="/Home" element={<Home />} />
                    {/* <Route path="/Upload" element={<UploadPage setUploadRedy/>} /> */}

                </Routes>
            </Router>
        </>
    )

}

export default RouterApp;