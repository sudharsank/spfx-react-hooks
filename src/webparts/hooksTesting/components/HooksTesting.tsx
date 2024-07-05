import * as React from 'react';
import { FC, useEffect, useState } from 'react';
import styles from './HooksTesting.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { useAppHelper } from '../../../useHelper';

export interface IHooksTestingProps {
    description: string;
    isDarkTheme: boolean;
    environmentMessage: string;
    hasTeamsContext: boolean;
    userDisplayName: string;
}

const HooksTesting: FC<IHooksTestingProps> = (props) => {
    const {
        description,
        isDarkTheme,
        environmentMessage,
        hasTeamsContext,
        userDisplayName
    } = props;
    const [loading, setLoading] = useState<boolean>(true);
    const {getListInfo, getListItems} = useAppHelper();

    const demo = async () => {
        setTimeout(() => {
            setLoading(false);
        }, 3000);

        const listInfo = await getListInfo('Announcements');
        console.log("List Information: ", listInfo);

        const listitems = await getListItems('Announcements');
        console.log("List items: ", listitems);
    };

    useEffect(() => {
        demo();
    }, []);

    return (
        <section className={`${styles.hooksTesting} ${hasTeamsContext ? styles.teams : ''}`}>
            <div className={styles.welcome}>
                <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
                <h2>Well done, {escape(userDisplayName)}! a very good start</h2>
                <div>{environmentMessage}</div>
                <div>Web part property value: <strong>{escape(description)}</strong></div>
            </div>
            <div>
                <div>
                    {loading ? "The component is loading" : "The component loaded"}
                </div>
                <h3>Welcome to SharePoint Framework!</h3>
                <p>
                    The SharePoint Framework (SPFx) is a extensibility model for Microsoft Viva, Microsoft Teams and SharePoint. It&#39;s the easiest way to extend Microsoft 365 with automatic Single Sign On, automatic hosting and industry standard tooling.
                </p>
                <h4>Learn more about SPFx development:</h4>
                <ul className={styles.links}>
                    <li><a href="https://aka.ms/spfx" target="_blank" rel="noreferrer">SharePoint Framework Overview</a></li>
                    <li><a href="https://aka.ms/spfx-yeoman-graph" target="_blank" rel="noreferrer">Use Microsoft Graph in your solution</a></li>
                    <li><a href="https://aka.ms/spfx-yeoman-teams" target="_blank" rel="noreferrer">Build for Microsoft Teams using SharePoint Framework</a></li>
                    <li><a href="https://aka.ms/spfx-yeoman-viva" target="_blank" rel="noreferrer">Build for Microsoft Viva Connections using SharePoint Framework</a></li>
                    <li><a href="https://aka.ms/spfx-yeoman-store" target="_blank" rel="noreferrer">Publish SharePoint Framework applications to the marketplace</a></li>
                    <li><a href="https://aka.ms/spfx-yeoman-api" target="_blank" rel="noreferrer">SharePoint Framework API reference</a></li>
                    <li><a href="https://aka.ms/m365pnp" target="_blank" rel="noreferrer">Microsoft 365 Developer Community</a></li>
                </ul>
            </div>
        </section>
    );
};

export default HooksTesting;
