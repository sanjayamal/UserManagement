import React from 'react';
import { Form, Input, Button, Switch } from 'antd';
import 'antd/dist/antd.css';
import { Row, Col } from 'antd';
import { createUser } from './GraphService';
import { config } from './Config';
import { UserAgentApplication } from 'msal';
import {useHistory  } from 'react-router-dom';


const AddUser = () => {
    let history = useHistory();

    const userAgentApplication = new UserAgentApplication({
        auth: {
            clientId: config.appId,
            redirectUri: config.redirectUri
        },
        cache: {
            cacheLocation: "sessionStorage",
            storeAuthStateInCookie: true
        }
    });

    const isInteractionRequired = (error: Error): boolean => {
        if (!error.message || error.message.length <= 0) {
            return false;
        }

        return (
            error.message.indexOf('consent_required') > -1 ||
            error.message.indexOf('interaction_required') > -1 ||
            error.message.indexOf('login_required') > -1
        );
    }

    const getAccessToken = async (scopes: string[]) => {

        try {
            var silentResult = await userAgentApplication.acquireTokenSilent({
                scopes: scopes
            });
            return silentResult.accessToken;
        } catch (err) {
            // If a silent request fails, it may be because the user needs
            // to login or grant consent to one or more of the requested scopes
            if (isInteractionRequired(err)) {
                var interactiveResult = await userAgentApplication.acquireTokenPopup({
                    scopes: scopes
                });

                return interactiveResult.accessToken;
            } else {
                throw err;
            }
        }
    }

    const passwordProfile = {
        "forceChangePasswordNextSignIn": true,
        "forceChangePasswordNextSignInWithMfa": false,
        "password": "Rajitha111#"
    }

    const getMailNickName = (email: string) => {
        let nickName = email.split('@');
        return nickName;
    }

    const onFinish = async (values: any) => {
        try {
            const accessToken = await getAccessToken(config.scopes);
            const { mail } = values;
            let mailNickname = getMailNickName(mail)[0]
            const requestBody = { ...values, mailNickname, passwordProfile: { ...passwordProfile },userPrincipalName:mail }
           
            console.log('access', accessToken);
            createUser(accessToken, requestBody)
            .then(res=>{
                console.log(res);
                history.push('/user')
            });
            
        } catch (error) {

            console.log(error);
        }

    };

    return (
        <>
            <Row>
                <Col span={6}></Col>
                <Col span={12}>
                    <Form
                        onFinish={onFinish}
                    >
                        <Form.Item
                            label='Account Enable'
                            name='accountEnabled'>
                            <Switch
                                checkedChildren="Yes" unCheckedChildren="No" defaultChecked={false} />
                        </Form.Item>

                        <Form.Item
                            label="Display Name"
                            name="displayName"
                            rules={[{ required: true, message: 'Please input your Display Name!' }]}
                        >
                            <Input />
                        </Form.Item>
                        <Form.Item
                            label="Email"
                            // name="userPrincipalName"
                            name='mail'
                            rules={[{ message: 'Please input your Email' }]}
                        >
                            <Input />
                        </Form.Item>
                        <br />
                        <Form.Item>
                            <Button type="primary" htmlType="submit">
                                Create
                            </Button>
                        </Form.Item>
                    </Form>
                </Col>
                <Col span={6}></Col>
            </Row>

        </>
    );
};

export default AddUser;