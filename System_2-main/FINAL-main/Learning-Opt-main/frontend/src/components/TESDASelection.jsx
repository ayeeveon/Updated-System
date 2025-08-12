import { useState } from "react";

function TESDASelection({ selected, onChange }) {
  const options = [
    { id: "tesda", label: "TESDA Certificate", desc: "- For TESDA Record of Candidate for Graduation" },
    { id: "custom", label: "Custom Template", desc: "- Create your custom template" },
  ];

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
      {/* Removed file upload / choose file UI */}
    </form>
  );
}

export default TESDASelection;