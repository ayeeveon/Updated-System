import { useState } from "react";

function Signatories(){
    
    const [signatory, setSignatory] = useState([{ sigName: "", sigTitle: "" }]);
    
    const MAX_SIGNATORIES = 4;

    const addSignatory = () => {
        if (signatory.length < MAX_SIGNATORIES) {
            setSignatory([...signatory, { sigName: "", sigTitle: "" }]);
        }
    };
    
    const handleChange = (idx, key, value) => {
        const updated = [...signatory];
        updated[idx][key] = value;
        setSignatory(updated);
    };
    
    const deleteSignatory = (idx) => {
        if (signatory.length === 1) return; // Prevent deleting the last field
        setSignatory(signatory.filter((_, i) => i !== idx));
    };

    return(
        <>
        <div className="flex flex-row justify-between items-end mt-3">
            <p className="text-md font-semibold">Signatories</p>
            <button type="button" onClick={addSignatory} className="text-xs font-semibold text-[#613d89] bg-[#a361ef90] px-5 py-0 rounded-md flex items-center gap-1"><p className="text-2xl transform-[translateY(-5%)]">+</p>Add Signatory</button>
        </div>

        <div className="flex flex-col gap-4 mt-2">
                {signatory.map((val, idx) => (
                    <div key={idx} className="flex flex-col gap-3 p-2 items-center rounded-md border-1">
                        <input
                            type="text"
                            className="mainInput h-1rem rounded-md p-1"
                            value={val.sigName}
                            onChange={e => handleChange(idx, "sigName", e.target.value)}
                            placeholder="Name"
                        />
                        <input
                            type="text"
                            className="mainInput h-1rem rounded-md p-1"
                            value={val.sigTitle}
                            onChange={e => handleChange(idx, "sigTitle", e.target.value)}
                            placeholder="Position/Title"
                        />

                        <div className="flex justify-between w-full ">
                            <button
                                type="button"
                                className="text-xs font-semibold text-[#613d89] bg-[#a361ef90] px-5 py-0 rounded-md flex items-center gap-1"
                            >
                                <p className="text-2xl transform-[translateY(-5%)]"></p>Upload
                            </button>
                            {signatory.length > 1 && (
                                <button
                                    type="button"
                                    className="text-xs bg-red-200 text-red-700 rounded px-2 py-1"
                                    onClick={() => deleteSignatory(idx)}
                                    title="Delete this field"
                                >
                                    Delete
                                </button>
                            )}
                        </div>
                    </div>
                ))}
            </div>
        </>
    )
}

export default Signatories