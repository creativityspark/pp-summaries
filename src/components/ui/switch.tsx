import {
  Switch as FluentSwitch,
  type SwitchProps as FluentSwitchProps,
} from "@fluentui/react-components";

import { cn } from "@/lib/utils";

type SwitchProps = Omit<FluentSwitchProps, "onChange" | "size"> & {
  size?: "sm" | "default";
  onCheckedChange?: (checked: boolean) => void;
};

function Switch({
  className,
  size = "default",
  onCheckedChange,
  ...props
}: SwitchProps) {
  return (
    <FluentSwitch
      data-slot="switch"
      className={cn(className)}
      size={size === "sm" ? "small" : "medium"}
      onChange={(_, data) => onCheckedChange?.(data.checked)}
      {...props}
    />
  );
}

export { Switch };
