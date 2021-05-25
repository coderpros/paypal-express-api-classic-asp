# Paypal Express API Implementation with JavaScript & Classic ASP

<a href="https://coderpro.net" target="_top"><img src="https://coderpro.net/media/1024/coderpro_logo_rounded_extra-90x90.webp" align="right" /></a>

[![Contributors][contributors-shield]][contributors-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]
[![MIT License][license-shield]][license-url]
[![LinkedIn][linkedin-shield]][linkedin-url]
[![Twitter](https://img.shields.io/twitter/url/https/twitter.com/cloudposse.svg?style=social&label=Follow%20%40coderProNet)](https://twitter.com/coderProNet)
[![GitHub](https://img.shields.io/github/followers/coderpros?label=Follow&style=social)](https://github.com/coderpros)

[contributors-shield]: https://img.shields.io/github/contributors/coderpros/paypal-express-api-classic-asp.svg?style=flat-square
[contributors-url]: https://github.com/coderpros/paypal-express-api-classic-asp/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/coderpros/paypal-express-api-classic-asp?style=flat-square
[forks-url]: https://github.com/coderpros/kickbox-asp/network/members
[stars-shield]: https://img.shields.io/github/stars/coderpros/paypal-express-api-classic-asp.svg?style=flat-square
[stars-url]: https://github.com/coderpros/kickbox-asp/stargazers
[issues-shield]: https://img.shields.io/github/issues/coderpros/paypal-express-api-classic-asp?style=flat-square
[issues-url]: https://github.com/coderpros/paypal-express-api-classic-asp/issues
[license-shield]: https://img.shields.io/github/license/coderpros/paypal-express-api-classic-asp?style=flat-square
[license-url]: https://github.com/coderpros/paypal-express-api-classic-asp/master/LICENSE
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=flat-square&logo=linkedin&colorB=555
[linkedin-url]: https://linkedin.com/company/coderpros
[twitter-shield]: https://img.shields.io/twitter/follow/coderpronet?style=social
[twitter-follow-url]: https://img.shields.io/twitter/follow/coderpronet?style=social
[github-shield]: https://img.shields.io/github/followers/coderpros?label=Follow&style=social
[github-follow-url]: https://img.shields.io/twitter/follow/coderpronet?style=social

> Sample Paypal Express API example with JavaScript and Classic ASP

## Notes
1) If you receive this error: 
msxml6.dll error '80072f0c'
A certificate is required to complete client authentication
/paypal.express.ui/inc/PaypalObject.class.asp, line 37

Thiat is because you do not have a *client* cert. 
The **second** answer on this StackOverflow question solves the problem: https://stackoverflow.com/questions/9212985/cant-use-https-with-serverxmlhttp-object

## Change Log
* 2021/05/11: Project creation / initial checkin.
