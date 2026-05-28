
import { Outlet } from "../types";
import { MASTER_OUTLETS } from "../data/masterOutlets";

export const getFilteredOutlets = (salesmanName: string, cycle: string): Outlet[] => {
  return MASTER_OUTLETS.filter(
    (outlet) => 
      outlet.salesman.toLowerCase().trim() === salesmanName.toLowerCase().trim() &&
      outlet.cycle.toLowerCase().trim() === cycle.toLowerCase().trim()
  );
};
