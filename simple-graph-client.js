const { Client } = require('@microsoft/microsoft-graph-client');

class SimpleGraphClient {
	constructor(token) {        
		if (!token || !token.trim()) {
			throw new Error('SimpleGraphClient: Invalid token received');
		}
		this.graphClient = Client.init({
			authProvider: (done) => {               
				done(null, token);
			}
		});
	}

	async sendMail(toAddress, subject, content){
		if(!toAddress || !toAddress.trim()) {
			throw new Error('SimpleGraphClient.sendMail(): Invalid `toAddress` parameter received.');
		}
		if (!subject || !subject.trim()) {
			throw new Error('SimpleGraphClient.sendMail(): Invalid `subject`  parameter received.');
		}
		if (!content || !content.trim()) {
			throw new Error('SimpleGraphClient.sendMail(): Invalid `content` parameter received.');
		}
		
		const mail = {
            body: {
                content: content, // `Hi there! I had this message sent from a bot. - Your friend, ${ graphData.displayName }!`,
                contentType: 'Text'
            },
            subject: subject, // `Message from a bot!`,
            toRecipients: [{
                emailAddress: {
                    address: toAddress
                }
            }]
	};
        return await this.graphClient
            .api('/me/sendMail')
            .post({ message: mail }, (error, res) => {
                if (error) {
                    throw error;
                } else {
                    return res;
                }
		});
}

async getRecentMail() {
        return await this.graphClient
            .api('/me/messages')
            .version('beta')
            .top(5)
            .get().then((res) => {
                return res;
            });
}

async getMe() {    
        return await this.graphClient
            .api('/me')
            .get().then((res) => {               
                return res;
            });
}

async getManager() {
        return await this.graphClient
            .api('/me/manager')
            .version('beta')            
            .get().then((res) => {               
                return res;
            });
}
}

module.exports.SimpleGraphClient = SimpleGraphClient;