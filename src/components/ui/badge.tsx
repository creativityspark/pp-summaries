import {
  Badge as FluentBadge,
  type BadgeProps as FluentBadgeProps,
} from "@fluentui/react-components";

import { cn } from "@/lib/utils";

type BadgeVariant =
  | "default"
  | "secondary"
  | "destructive"
  | "outline"
  | "ghost"
  | "link";

const appearanceByVariant: Record<
  BadgeVariant,
  FluentBadgeProps["appearance"]
> = {
  default: "filled",
  secondary: "tint",
  destructive: "tint",
  outline: "outline",
  ghost: "ghost",
  link: "ghost",
};

type BadgeProps = Omit<FluentBadgeProps, "appearance" | "color" | "shape"> & {
  variant?: BadgeVariant;
};

function Badge({ className, variant = "default", ...props }: BadgeProps) {
  return (
    <FluentBadge
      data-slot="badge"
      data-variant={variant}
      appearance={appearanceByVariant[variant]}
      color={variant === "destructive" ? "danger" : "brand"}
      shape="rounded"
      className={cn(className)}
      {...props}
    />
  );
}

export { Badge };
