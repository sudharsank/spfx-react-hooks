import { SPFI } from "@pnp/sp"
import { getSP } from "./pnp.config"
import { useCallback } from "react";

export const useAppHelper = () => {
    let _sp: SPFI = getSP();

    const getListInfo = useCallback(async (listname: string): Promise<any> => {
        try {
            return await _sp.web.lists.getByTitle(listname)();
        } catch (err) {
            console.log("getListInfo: ", err);
        }
    }, [_sp]);

    const getListItems = useCallback(async (listname: string): Promise<any> => {
        try {
            return await _sp.web.lists.getByTitle(listname).items();
        } catch (err) {
            console.log("Get Item Info: ", err);
        }
    }, [_sp]);

    return {
        getListInfo,
        getListItems
    }
}