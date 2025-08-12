import { useState } from "react";

function CertificateDetails() {
    // Each field group is an object with two values
    const [fields, setFields] = useState([{ field1: "", field2: "" }]);

    const addField = () => {
        setFields([...fields, { field1: "", field2: "" }]);
    };

    const handleChange = (idx, key, value) => {
        const updated = [...fields];
        updated[idx][key] = value;
        setFields(updated);
    };

    const deleteField = (idx) => {
        if (fields.length === 1) return; // Prevent deleting the last field
        setFields(fields.filter((_, i) => i !== idx));
    };

    return (
        <>
            {/* <p className="text-md font-semibold">Certificate Title</p>
            <input type="text" id="certTitle" className="mainInput h-1rem rounded-md p-1 my-2" placeholder="e.g., Certificate of Completion"/> */}

            <div className="flex flex-row justify-between items-end mt-3">
                <p className="text-md font-semibold">Certificate Details Fields</p>
                {/* <button type="button" onClick={addField} className="text-xs font-semibold text-[#613d89] bg-[#a361ef90] px-5 rounded-md flex items-center gap-1"><p className="text-2xl transform-[translateY(-5%)]">+</p>Add field</button> */}
            </div>

            <div className="flex flex-col gap-1">
                {fields.map((val, idx) => (
                    <div key={idx} className="certDetails flex gap-2 bg-[#ffffff1a] p-2 mt-2 items-center rounded-md">
                        {/* <input
                            type="text"
                            className="mainInput h-1rem rounded-md p-1 my-2"
                            value={val.field1}
                            onChange={e => handleChange(idx, "field1", e.target.value)}
                            placeholder="Field Label"
                        /> */}
                        <input
                            type="text"
                            className="mainInput h-1rem rounded-md p-1 my-2"
                            value={val.field2}
                            onChange={e => handleChange(idx, "field2", e.target.value)}
                            placeholder="Description"
                        />
                        {fields.length > 1 && (
                            <button
                                type="button"
                                className="text-xs bg-red-200 text-red-700 rounded px-2 py-1"
                                onClick={() => deleteField(idx)}
                                title="Delete this field"
                            >
                                Delete
                            </button>
                        )}
                    </div>
                ))}
            </div>
        </>
    );
}

export default CertificateDetails;