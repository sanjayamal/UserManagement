import React, { useState, useEffect } from 'react';
import { config } from './Config';
import { UserAgentApplication } from 'msal';
import { getUser, deleteUser, updateUser, addUserGroup, getGroup, getMemberGroups, deleteUserGroup } from './GraphService';
import { Table, Button, Drawer, Form, Input, Modal, Select } from 'antd';
import 'antd/dist/antd.css';
import { DeleteOutlined } from '@ant-design/icons';
import { constants } from 'buffer';

const User = () => {

    const [users, setUser] = useState<any>([]);
    const [drawerVisible, setDrawerVisible] = useState<boolean>(false)
    const [user, setOneUser] = useState<any>();
    const [userId, setUserId] = useState<any>();
    const [groups, setGroups] = useState<any>([]);
    const [groupId, setGroupId] = useState<any>();
    const [visibleModel, setvisibleModel] = useState<boolean>();
    const [userGroups, setUserGroups] = useState<any>([]);

    const { confirm } = Modal;
    const { Option } = Select;

    useEffect(() => {
        getUserService();
        getGroupToModel();

    }, [user]);



    const columns = [
        {
            title: 'Display Name',
            dataIndex: 'displayName',
            key: 'displayName',
        },
        {
            title: 'Mail',
            dataIndex: 'mail',
            key: 'mail',
        },
        {
            title: '',
            dataIndex: 'id',
            key: 'id',
            render: (key: any, test: any) => <>
                <Button style={{ 'marginRight': '3px' }} onClick={() => userUpdate(test)}>Update</Button>
                <Button style={{ 'marginRight': '3px' }} onClick={() => { UserGroup(key) }}>Group</Button>
                <Button onClick={() => { userDelete(key) }} >Delete</Button>

            </>,
        },

    ];

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

    const getUserService = async () => {
        const accessToken = await getAccessToken(config.scopes);
        console.log(accessToken)

        getUser(accessToken)
            .then((result: any) => {
                const { value } = result;
                setUser(value)
                console.log(value)
            });

    }

    const getGroupToModel = async () => {
        const accessToken = await getAccessToken(config.scopes);
        getGroup(accessToken)
            .then((result: any) => {
                const { value } = result;
                setGroups(value)
                console.log(value)
            });
    }
    const UserGroup = async (id: any) => {
        console.log(groups)
        const accessToken = await getAccessToken(config.scopes);
        setUserId(id);
        setvisibleModel(true);
        getMemberGroups(accessToken, id, reqBody)
            .then((result: any) => {
                const { value } = result
                console.log(result)
                console.log(value)
                setUserGroups(value);
            })
    }

    const group = {
        'members@odata.bind': [
            `https://graph.microsoft.com/v1.0/directoryObjects/${userId}`
        ]

    };
    const reqBody = {
        securityEnabledOnly: true
    };

    const userAddGroup = async () => {
        const accessToken = await getAccessToken(config.scopes);
        addUserGroup(accessToken, groupId, group)
            .then((result: any) => {
                setvisibleModel(false);
            });
    }
    const userDelete = async (id: any) => {

        confirm({
            title: 'Do you Want to delete these items?',
            // content: 'Some descriptions',
            async onOk() {
                const data = users.filter((user: any) => user.id !== id)
                setUser(data)
                const accessToken = await getAccessToken(config.scopes);
                deleteUser(accessToken, id)
                    .then(res => console.log(res))
                    .catch(res => console.log(res));
            },
            onCancel() {
                console.log('Cancel');
            },
        });
        // userAddGroup();

    }

    const deleteMember = async (gId:any) => {
        const accessToken = await getAccessToken(config.scopes);
        deleteUserGroup(accessToken, gId, userId)
            .then((result: any) => {
                console.log(result)
                setvisibleModel(false);
            });
       // console.log("delete member function",groupId)
    }
    const userUpdate = (user: any) => {
        setDrawerVisible(true);
        setOneUser(user)

    }


    const userUpdateSave = async (values: any) => {

        const accessToken = await getAccessToken(config.scopes);
        values.id = values.id !== undefined ? values.id : (user.id !== null ? user.id : '');
        values.displayName = values.displayName !== undefined || null ? values.displayName : (user.displayName !== null ? user.displayName : '');
        values.givenName = values.givenName !== undefined || null ? values.givenName : (user.givenName !== null ? user.givenName : null);
        values.jobTitle = values.jobTitle !== undefined || null ? values.jobTitle : (user.jobTitle !== null ? user.jobTitle : null);
        values.mobilePhone = values.mobilePhone !== undefined || null ? values.mobilePhone : (user.mobilePhone !== null ? user.mobilePhone : null);
        values.officeLocation = values.officeLocation !== undefined || null ? values.officeLocation : (user.officeLocation !== null ? user.officeLocation : null);
        values.surname = values.surname !== undefined || null ? values.surname : (user.surname !== null ? user.surname : null);
        values.userPrincipalName = values.userPrincipalName !== undefined || null ? values.userPrincipalName : (user.userPrincipalName !== null ? user.userPrincipalName : null);
        values.mail = values.mail !== undefined || null ? values.mail : (user.mail !== null ? user.mail : null);


        updateUser(accessToken, values)
            .then((res: any) => {
                setDrawerVisible(false);
                setOneUser(values)
            })
    }

    const onClose = () => {
        setDrawerVisible(false)
    }
    const handleCancel = () => {
        setvisibleModel(false)
    }

    const onChangeGroup = (groupId: any) => {
        setGroupId(groupId);
        console.log("onChange ",groupId)
    }

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

    // const options:any = [];
    // groups?.map((item: any) => {
    //     const value = item.displayName;
    //     options.push({
    //         value
    //     });
    // })


    return (
        <>
            <Button href='/adduser' >Create New User</Button>
            <br />
            <Table dataSource={users} columns={columns} style={{ 'paddingTop': '10px' }} />

            <Drawer
                title="Create a new account"
                width={720}
                onClose={onClose}
                visible={drawerVisible}
                bodyStyle={{ paddingBottom: 80 }}

            >
                <Form
                    onFinish={userUpdateSave}
                    initialValues={{
                        'displayName': user === undefined ? '' : (user.displayName !== null ? user.displayName : ''),
                        'givenName': user === undefined ? '' : (user.givenName !== null ? user.givenName : ''),
                        'id': user === undefined ? '' : (user.id !== null ? user.id : ''),
                        'jobTitle': user === undefined ? '' : (user.jobTitle !== null ? user.jobTitle : ''),
                        'mobilePhone': user === undefined ? '' : (user.mobilePhone !== null ? user.mobilePhone : ''),
                        'officeLocation': user === undefined ? '' : (user.officeLocation !== null ? user.officeLocation : ''),
                        'surname': user === undefined ? '' : (user.surname !== null ? user.surname : ''),
                        'userPrincipalName': user === undefined ? '' : (user.userPrincipalName !== null ? user.userPrincipalName : ''),
                        'mail': user === undefined ? '' : (user.mail !== null ? user.mail : ''),
                    }}
                >


                    <Form.Item
                        label="Display Name"
                        name="displayName"
                    >
                        <Input />
                    </Form.Item>
                    <Form.Item
                        label="Email"
                        name="mail"
                    >
                        <Input />
                    </Form.Item>
                    <br />
                    <Form.Item>
                        <Button type="primary" htmlType="submit">
                            update
                            </Button>
                        <Button onClick={onClose} type="default" htmlType="submit">
                            Cancel
                            </Button>
                    </Form.Item>
                </Form>
            </Drawer>

            <Modal
                visible={visibleModel}
                title="Select Groups"
                onOk={userAddGroup}
                onCancel={handleCancel}
                footer={[
                    <Button key="back" onClick={handleCancel}>
                        Return
            </Button>,
                    <Button key="submit" type="primary" onClick={userAddGroup}>
                        Submit
            </Button>,
                ]}
            >
                <Select
                    showSearch
                    style={{ width: 200 }}
                    placeholder="Select a Group"
                    onChange={onChangeGroup}
                >
                    {groups?.map((item: any) =>
                        <Option key={item.id} value={item.id}>{item.displayName}</Option>

                    )}
                </Select>

                {groups?.filter((item: any) => userGroups.includes(item.id))
                    .map((item: any) =>
                        <div style={{ marginTop: '5px' }}>
                            {item.displayName}
                            <Button onClick={()=>deleteMember(item.id)} type="primary" shape="circle" icon={<DeleteOutlined />} size="small" style={{ marginLeft: '10px' }}>
                            </Button>
                        </div>
                    )}
            </Modal>

        </>
    );



}

export default User;