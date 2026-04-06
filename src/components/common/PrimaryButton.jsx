export default function PrimaryButton({
  onClick,
  children,
  tone = "primary",
  type = "button",
  disabled = false,
  style,
}) {
  return (
    <button
      className={`btn btn--${tone}`}
      onClick={onClick}
      type={type}
      disabled={disabled}
      style={style}
    >
      {children}
    </button>
  );
}
