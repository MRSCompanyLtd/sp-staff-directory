import * as React from "react";
import { createElement, useEffect, useRef, useState } from "react";
import { Log } from "@microsoft/sp-core-library";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { ILivePersonaProps } from "./ILivePersonaProps";
import styles from "./Person.module.scss";

const LIVE_PERSONA_COMPONENT_ID: string =
    "914330ee-2df2-4f6e-a858-30c23a812408";

const LivePersona: React.FunctionComponent<ILivePersonaProps> = ({
    upn, template, disableHover, serviceScope, children
}) => {
    const [isComponentLoaded, setIsComponentLoaded] = useState<boolean>(false);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sharedLibrary: any = useRef<any>();

    useEffect(() => {
        (async () => {
            if (!isComponentLoaded) {
                try {
                    sharedLibrary.current =
                        await SPComponentLoader.loadComponentById(
                            LIVE_PERSONA_COMPONENT_ID
                        );
                    setIsComponentLoaded(true);
                } catch (error) {
                    Log.error(`[LivePersona]`, error, serviceScope);
                }
            }
        })()
            .then(() => {
                /* no-op; */
            })
            .catch(() => {
                /* no-op; */
            });
    // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    let renderPersona: JSX.Element = null;
    if (isComponentLoaded) {
        renderPersona = createElement(
            sharedLibrary.current.LivePersonaCard,
            {
                className: styles.person,
                clientScenario: "livePersonaCard",
                disableHover: disableHover,
                hostAppPersonaInfo: {
                    PersonaType: "User",
                },
                upn: upn,
                legacyUpn: upn,
                serviceScope: serviceScope,
            },
            createElement("div", {}, template)
        );
    }
    return renderPersona;
};

export default LivePersona;
