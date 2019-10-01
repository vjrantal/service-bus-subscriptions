// Imports of official SDKs
import { ServiceClientCredentials, WebResource } from "@azure/ms-rest-js";
import { ServiceBusManagementClient } from "@azure/arm-servicebus";
import { ServiceBusClient, ReceiveMode, ServiceBusMessage, MessagingError, TopicClient, Sender, SendableMessageInfo, SubscriptionClient, Receiver } from "@azure/service-bus";

// Imports related to injecting credentials from the app to the SDKs
import { MsalAuthProvider } from 'react-aad-msal';
import { TokenType, TokenProvider, TokenInfo } from '@azure/amqp-common';
import { EventEmitter } from "events";

const subscriptionId = process.env.REACT_APP_SUBSCRIPTION_ID || '';
const resourceGroupName = process.env.REACT_APP_RESOURCE_GROUP_NAME || '';
const namespaceName = process.env.REACT_APP_NAMESPACE_NAME || '';
const topicName = process.env.REACT_APP_TOPIC_NAME || '';
const subscriptionName = process.env.REACT_APP_SUBSCRIPTION_NAME || '';

class ReactAADCredentials implements ServiceClientCredentials {
    provider: MsalAuthProvider;

    constructor(provider: MsalAuthProvider) {
        this.provider = provider;
    }

    async signRequest(webResource: WebResource): Promise<WebResource> {
        const token = await this.provider.getAccessToken({
            scopes: ['https://management.azure.com/user_impersonation']
        });
        webResource.headers.set('Authorization', 'Bearer ' + token.accessToken);
        return Promise.resolve(webResource);
    }
}

class ReactAADTokenInfo implements TokenInfo {
    tokenType: TokenType;
    token: string;
    expiry: number;

    constructor(tokenType: TokenType, token: string, expiry: number) {
        this.tokenType = tokenType;
        this.token = token;
        this.expiry = expiry;
    }
}

class ReactAADTokenProvider implements TokenProvider {
    provider: MsalAuthProvider;
    tokenRenewalMarginInSeconds: number;
    tokenValidTimeInSeconds: number;

    constructor(provider: MsalAuthProvider) {
        this.provider = provider;
        this.tokenRenewalMarginInSeconds = 900;
        this.tokenValidTimeInSeconds = 3600;
    }

    async getToken(audience?: string): Promise<TokenInfo> {
        const token = await this.provider.getAccessToken({
            scopes: ['https://servicebus.azure.net/user_impersonation']
        });
        const expiry = (new Date(token.expiresOn).getTime() - new Date().getTime()) / 1000;
        const tokenInfo = new ReactAADTokenInfo(TokenType.CbsTokenTypeJwt, token.accessToken, expiry);
        return Promise.resolve(tokenInfo);
    }
}

export class ServiceBus extends EventEmitter {
    provider?: MsalAuthProvider;
    serviceBusClient?: ServiceBusClient;
    serviceBusManagementClient?: ServiceBusManagementClient;
    topicClient?: TopicClient;
    topicSender?: Sender;
    subscriptionClient?: SubscriptionClient;
    subscriptionReceiver?: Receiver;

    async initialize(provider: MsalAuthProvider) {
        this.provider = provider;

        const creds = new ReactAADCredentials(provider);
        this.serviceBusManagementClient = new ServiceBusManagementClient(creds, subscriptionId);

        const tokenProvider = new ReactAADTokenProvider(provider);
        this.serviceBusClient = ServiceBusClient.createFromTokenProvider(namespaceName + '.servicebus.windows.net', tokenProvider);

        this.topicClient = this.serviceBusClient.createTopicClient(topicName);
        this.topicSender = this.topicClient.createSender();

        await this.createSubscription();
        this.subscriptionClient = this.serviceBusClient.createSubscriptionClient(topicName, subscriptionName);
        this.subscriptionReceiver = this.subscriptionClient.createReceiver(ReceiveMode.peekLock);

        this.subscriptionReceiver.registerMessageHandler((brokeredMessage: ServiceBusMessage) => {
            const delay = new Date().getTime() - new Date(brokeredMessage.body).getTime();
            this.emit('result', `Received message with body: ${brokeredMessage.body} with delay: ${delay}ms`);
            return Promise.resolve();
        }, (error: MessagingError | Error) => {
            this.emit('result', "Error occurred: " + error);
        });
    }

    async uninitialize() {
        if (!this.serviceBusClient) {
            return
        }
        await this.serviceBusClient.close();
    }

    async send() {
        const sendableMessage: SendableMessageInfo = {
            body: new Date().toISOString()
        };
        if (this.topicSender) {
            await this.topicSender.send(sendableMessage);
        } else {
            this.emit('result', 'No sender available');
        }
    }

    async createSubscription() {
        if (!this.serviceBusManagementClient) {
            return;
        }
        const parameters = {
            autoDeleteOnIdle: 'PT5M'
        };
        this.serviceBusManagementClient.subscriptions.createOrUpdate(resourceGroupName, namespaceName, topicName, subscriptionName, parameters).then((result) => {
            this.emit('result', JSON.stringify(result) + '');
        }).catch((error) => {
            this.emit('result', error + '');
        });
    }

    async getSubscriptions() {
        if (!this.serviceBusManagementClient) {
            return;
        }
        this.serviceBusManagementClient.subscriptions.listByTopic(resourceGroupName, namespaceName, topicName).then((result) => {
            this.emit('result', JSON.stringify(result) + '');
        }).catch((error) => {
            this.emit('result', error + '');
        });
    }
}
