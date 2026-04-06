export default function QtyInput({
  value,
  onChange,
  min = 0,
  placeholder = "",
  className = "qtyInput",
}) {
  return (
    <input
      type="text"
      inputMode="numeric"
      pattern="[0-9]*"
      className={className}
      placeholder={placeholder}
      value={value === 0 ? "" : String(value)}
      onChange={(e) => {
        const v = e.target.value;

        // ✅ tillåt tomt medan man skriver
        if (v === "") {
          onChange(0);
          return;
        }

        // ✅ tillåt bara siffror
        if (!/^\d+$/.test(v)) return;

        const n = Number(v);
        if (n < min) return;

        onChange(n);
      }}
    />
  );
}