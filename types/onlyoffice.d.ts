declare global {
    interface Window {
        /**
        * OnlyOffice plugin globals - intentionally untyped for simplicity
        * Use (window.Asc as any) in code
        */
        Asc?: any;
    }

    /**
    * Document Builder API - available inside Asc.plugin.callCommand()
    * Use (Api as any) in code
    */
    const Asc: any;
    const Api: any;
}

export {};
