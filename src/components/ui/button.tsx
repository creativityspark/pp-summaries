import * as React from "react";
import {
  Button as FluentButton,
  type ButtonProps as FluentButtonProps,
} from "@fluentui/react-components";

import { cn } from "@/lib/utils";

type ButtonVariant =
  | "default"
  | "outline"
  | "secondary"
  | "ghost"
  | "destructive"
  | "link";
type ButtonSize =
  | "default"
  | "xs"
  | "sm"
  | "lg"
  | "icon"
  | "icon-xs"
  | "icon-sm"
  | "icon-lg";

const appearanceByVariant: Record<
  ButtonVariant,
  FluentButtonProps["appearance"]
> = {
  default: "primary",
  outline: "secondary",
  secondary: "secondary",
  ghost: "subtle",
  destructive: "primary",
  link: "transparent",
};

const fluentSizeBySize: Record<ButtonSize, FluentButtonProps["size"]> = {
  default: "medium",
  xs: "small",
  sm: "small",
  lg: "large",
  icon: "medium",
  "icon-xs": "small",
  "icon-sm": "small",
  "icon-lg": "large",
};

type ButtonProps = React.ComponentProps<"button"> & {
  variant?: ButtonVariant;
  size?: ButtonSize;
};

function Button({
  className,
  variant = "default",
  size = "default",
  ...props
}: ButtonProps) {
  return (
    <FluentButton
      data-slot="button"
      data-variant={variant}
      data-size={size}
      appearance={appearanceByVariant[variant]}
      size={fluentSizeBySize[size]}
      shape="rounded"
      className={cn(
        variant === "destructive" && "fluent-destructive",
        className,
      )}
      {...props}
    />
  );
}

export { Button };
