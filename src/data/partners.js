// Данные партнёров. Единственное место, где живут имена, фото и контакты.
// Чтобы добавить партнёра: добавь запись, положи фото в public/ и пересобери сайт.

export const partners = {
  'chuguevskaya-galina': {
    slug: 'chuguevskaya-galina',
    fullName: 'Чугуевская Галина Викторовна',
    shortName: 'Галина',
    role: 'VIP-партнёр',
    city: 'Геленджик',
    photo: '/gv.webp',
    ogImage: '/og/galina.jpg',
    title: 'Чугуевская Галина — VIP-партнёр Greenleaf в Геленджике',
    description:
      'Чугуевская Галина Викторовна — VIP-партнёр Greenleaf в Геленджике. Премиальный уход, витамины, товары для дома и одежды. Помощь с выбором и оформлением заказа.',
    contacts: {
      phones: ['+79168692616', '+79384021767'],
      max: 'https://max.ru/u/f9LHodD0cOKzWN-GBIeaeUHPu68-AhiRXjm3YqvnhzArIvVZcyx3sfmzN7s',
      telegram: 'modagelendzhik',
    },
  },

  'toloknova-natalia': {
    slug: 'toloknova-natalia',
    fullName: 'Толокнова Наталия Валерьевна',
    shortName: 'Наталия',
    role: 'VIP-партнёр',
    city: 'Геленджик',
    photo: '/i.webp',
    ogImage: '/og/natalia.jpg',
    title: 'Толокнова Наталия — VIP-партнёр Greenleaf в Геленджике',
    description:
      'Толокнова Наталия Валерьевна — VIP-партнёр Greenleaf в Геленджике. Премиальный уход, витамины, товары для дома и одежды. Помощь с выбором и оформлением заказа.',
    contacts: {
      phones: ['+79384289200'],
      max: 'https://max.ru/u/f9LHodD0cOKb4Eyczebr3LMVLosYcO9pub7ipqg8pGUVWp6DB4G-sq9B6c0',
      vk: 'https://vk.ru/vsya_v_myle',
      email: 'natalytoloknova1978@gmail.com',
    },
  },
};

export const partnerList = Object.values(partners);

export const defaultPartner = partners['chuguevskaya-galina'];

export const getPartner = (slug) => partners[slug];

export const contactsHrefOf = (partner) => `/partners/${partner.slug}/contacts/`;

export const formatPhone = (raw) => {
  const digits = String(raw).replace(/\D/g, '');
  if (digits.length === 11) {
    return `+${digits[0]} (${digits.slice(1, 4)}) ${digits.slice(4, 7)}-${digits.slice(7, 9)}-${digits.slice(9, 11)}`;
  }
  return raw;
};
