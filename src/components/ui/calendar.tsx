"use client"

import * as React from "react"
import { DayPicker } from "react-day-picker"
import { ChevronLeft, ChevronRight } from "lucide-react"
import { cn } from "@/lib/utils"

export type CalendarProps = React.ComponentProps<typeof DayPicker>

function Calendar({ className, classNames, showOutsideDays = true, ...props }: CalendarProps) {
  return (
    <DayPicker
      showOutsideDays={showOutsideDays}
      className={cn("p-3", className)}
      classNames={{
        months: "flex flex-col sm:flex-row gap-4",
        month: "space-y-3",
        month_caption: "flex justify-center pt-1 relative items-center h-8",
        caption_label: "text-sm font-medium",
        nav: "flex items-center gap-1",
        button_previous: cn(
          "absolute left-1 h-7 w-7 inline-flex items-center justify-center rounded-md border border-input bg-transparent opacity-60 hover:opacity-100 hover:bg-accent transition-opacity"
        ),
        button_next: cn(
          "absolute right-1 h-7 w-7 inline-flex items-center justify-center rounded-md border border-input bg-transparent opacity-60 hover:opacity-100 hover:bg-accent transition-opacity"
        ),
        month_grid: "w-full border-collapse",
        weekdays: "flex",
        weekday: "text-muted-foreground w-9 font-normal text-[0.8rem] flex items-center justify-center",
        week: "flex w-full mt-1",
        day: "relative p-0 text-center",
        day_button: cn(
          "h-9 w-9 p-0 font-normal text-sm inline-flex items-center justify-center rounded-md",
          "hover:bg-accent hover:text-accent-foreground focus:outline-none focus-visible:ring-2 focus-visible:ring-ring"
        ),
        selected: "[&>button]:bg-primary [&>button]:text-primary-foreground [&>button]:hover:bg-primary [&>button]:hover:text-primary-foreground",
        today: "[&>button]:bg-accent [&>button]:text-accent-foreground",
        outside: "[&>button]:text-muted-foreground [&>button]:opacity-50",
        disabled: "[&>button]:text-muted-foreground [&>button]:opacity-30 [&>button]:pointer-events-none",
        range_start: "[&>button]:bg-primary [&>button]:text-primary-foreground rounded-l-md [&>button]:rounded-r-none",
        range_end: "[&>button]:bg-primary [&>button]:text-primary-foreground rounded-r-md [&>button]:rounded-l-none",
        range_middle: "bg-accent/40 [&>button]:bg-transparent [&>button]:rounded-none [&>button]:hover:bg-transparent",
        hidden: "invisible",
        ...classNames,
      }}
      components={{
        Chevron: ({ orientation }) =>
          orientation === "left" ? (
            <ChevronLeft className="h-4 w-4" />
          ) : (
            <ChevronRight className="h-4 w-4" />
          ),
      }}
      {...props}
    />
  )
}
Calendar.displayName = "Calendar"

export { Calendar }
