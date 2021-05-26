const resumeResolvers = require('./resume');
const usersResolvers = require('./users');
const usersEducation = require('./educations');
const miscResolvers = require('./misc');

module.exports = {
    Query : {
        ...resumeResolvers.Query
    },
    Mutation: {
        ...usersResolvers.Mutation,
        ...resumeResolvers.Mutation,
        ...usersEducation.Mutation,
        ...miscResolvers.Mutation
    }
}