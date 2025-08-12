import { useEffect, useState } from "react";

function TemplateSelection({ selected, onChange }) {
    const options = [
        { id: "ojt", label: "OJT Certificate", desc: "- For OJT completion" },
        { id: "immersion", label: "Work Immersion Certificate", desc: "- For work immersion completion" },
        { id: "custom", label: "Custom Template", desc: "- Create your custom template" },
    ];

    const [fileMap, setFileMap] = useState({
        ojt: null,
        immersion: null,
        custom: null,
    });

    useEffect(() => {
        fetch("http://localhost:5000/api/templates")
            .then(res => res.json())
            .then(data => {
                const updated = { ojt: null, immersion: null, custom: null };
                if (data.ojt) {
                    updated.ojt = {
                        name: data.ojt.name,
                        isDefault: true,
                        url: `http://localhost:5000${data.ojt.url}`,
                    };
                }
                if (data.immersion) {
                    updated.immersion = {
                        name: data.immersion.name,
                        isDefault: true,
                        url: `http://localhost:5000${data.immersion.url}`,
                    };
                }
                setFileMap(updated);
            });
    }, []);

    const handleFileChange = async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append("template", file);
        formData.append("type", selected);

        try {
            const res = await fetch("http://localhost:5000/api/upload-template", {
                method: "POST",
                body: formData,
            });

            if (res.ok) {
                setFileMap((prev) => ({
                    ...prev,
                    [selected]: {
                        name: file.name,
                        isDefault: false,
                    },
                }));
                alert("Upload successful!");
            } else {
                alert("Upload failed.");
            }
        } catch (err) {
            console.error("Upload error:", err);
        }
    };

    const handleRemoveFile = () => {
        setFileMap((prev) => ({ ...prev, [selected]: null }));
    };

    const currentFile = fileMap[selected];

    return (
        <form className="flex flex-col gap-4">
            {/* Template options */}
            {options.map((opt) => (
                <div
                    key={opt.id}
                    className={`radioSelect ${selected === opt.id ? "radioSelect--selected" : ""}`}
                >
                    <input
                        type="radio"
                        name="template"
                        id={opt.id}
                        value={opt.id}
                        className="accent-[#a361ef]"
                        checked={selected === opt.id}
                        onChange={() => onChange(opt.id)}
                    />
                    <label htmlFor={opt.id} className="flex flex-col">
                        <h3 className="text-xl">{opt.label}</h3>
                        <p className="text-xs">{opt.desc}</p>
                    </label>
                </div>
            ))}

            {/* File upload or preview */}
            <div className="mt-4">
                {currentFile ? (
                    <div className="flex items-center gap-2">
                        <span className="text-sm font-medium">{currentFile.name}</span>
                        <button
                            type="button"
                            className="text-xs font-semibold text-red-700 bg-red-200 px-3 py-1 rounded-md hover:bg-red-300"
                            onClick={handleRemoveFile}
                        >
                            Remove
                        </button>
                    </div>
                ) : (
                    <div className="flex items-center gap-2">
                        <label
                            htmlFor={`file-upload-${selected}`}
                            className="text-l font-semibold text-[#613d89] bg-[#a361ef90] px-3 py-2 rounded-md cursor-pointer hover:bg-[#a361ef] inline-block"
                        >
                            Choose File
                        </label>
                        <input
                            id={`file-upload-${selected}`}
                            type="file"
                            accept=".ppt,.pptx"
                            onChange={handleFileChange}
                            className="hidden"
                        />
                    </div>
                )}
            </div>
        </form>
    );
}

export default TemplateSelection;