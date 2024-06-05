import type { BryntumSchedulerProps } from "@bryntum/scheduler-react";

const schedulerConfig: Partial<BryntumSchedulerProps> = {
  features: {
    stripe: true,
    dependencies: true,
  },
  columns: [{ text: "Name", field: "name", width: 250 }],
  viewPreset: "hourAndDay",
  startDate: new Date(2024, 6, 1, 9),
  endDate: new Date(2024, 6, 1, 17),
  barMargin: 10,
  selectionMode: {
    multiSelect: false,
  },
};

export { schedulerConfig };
