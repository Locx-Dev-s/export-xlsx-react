import * as React from "react";

import CreateSheet from './components/CreateSheet/CreateSheet'

const App = () => {
    return (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10, padding: "30px" }}>
            <CreateSheet />
        </div>
    );
};

export default App;
